package com.lannstark.excel.sxssf;

import com.lannstark.excel.ExcelFile;
import com.lannstark.exception.ExcelInternalException;
import com.lannstark.resource.*;
import lombok.Getter;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.*;
import java.util.stream.Collectors;

import static com.lannstark.utils.SuperClassReflectionUtils.getField;

/**
 * SXSSFExcelFile 클래스는 Apache POI의 SXSSF(SXSSFWorkbook)를 이용하여 Excel 파일을 생성하고 데이터의 렌더링을 지원하는 추상 클래스입니다.
 *  - 템플릿 메서드 패턴 기반
 * @param <T> 렌더링할 데이터 타입
 */
public abstract class SXSSFExcelFile<T> implements ExcelFile<T> {

	protected static final SpreadsheetVersion supplyExcelVersion = SpreadsheetVersion.EXCEL2007;
	private static final int COLUMN_WIDTH_PADDING = 512;

    // List 구분자 설정
    // 기본값: 쉼표+공백
	@Getter
    private String listSeparator = ", ";

	protected SXSSFWorkbook wb;
	protected Sheet sheet;
	protected ExcelRenderResource resource;

	/**
	 *SXSSFExcelFile
	 * @param type Class type to be rendered
	 */
	public SXSSFExcelFile(Class<T> type) {
		this(Collections.emptyList(), type, new DefaultDataFormatDecider());
	}

	/**
	 * SXSSFExcelFile
	 * @param data List Data to render an Excel file. Data should have at least one @ExcelColumn on fields
	 * @param type Class type to be rendered
	 */
	public SXSSFExcelFile(List<T> data, Class<T> type) {
		this(data, type, new DefaultDataFormatDecider());
	}

	/**
	 * SXSSFExcelFile
	 * @param data List Data to render an Excel file. Data should have at least one @ExcelColumn on fields
	 * @param type Class type to be rendered
	 * @param dataFormatDecider Custom DataFormatDecider
	 */
	public SXSSFExcelFile(List<T> data, Class<T> type, DataFormatDecider dataFormatDecider) {
		validateData(data);
		this.wb = new SXSSFWorkbook();
		this.resource = ExcelRenderResourceFactory.prepareRenderResource(type, wb, dataFormatDecider);
		renderExcel(data);
	}

    /**
     * 데이터를 유효성 검증합니다.
     * 후크 메서드 - 구체적인 구현을 하위 클래스에 위임
     * @param data 유효성을 검증할 데이터 리스트
     */
	protected void validateData(List<T> data) {}

    /**
     * List 값을 문자열로 변환할 때 사용할 구분자를 설정합니다.
     * @param separator 구분자 (예: ", ", "; ", "\n" 등)
     */
    public void setListSeparator(String separator) {
        this.listSeparator = separator != null ? separator : ", ";
    }

    /**
     * 제공된 데이터를 바탕으로 Excel 파일을 렌더링합니다.
     *
     * @param data Excel 파일에 렌더링할 데이터 리스트
     */
	protected abstract void renderExcel(List<T> data);

    /**
     * 새 시트에 헤더를 생성하고 렌더링합니다.
     * 주어진 시트 객체를 기반으로 헤더를 생성하며, 해당 헤더는 ExcelHeader와 매핑된 필드 경로를 사용하여 생성됩니다.
     * 헤더의 셀 병합 및 스타일 지정 작업도 이 메서드에서 수행됩니다.
     *
     * @param sheet 헤더를 생성할 대상 시트
     * @param rowIndex 시작 행 인덱스
     * @param columnStartIndex 시작 열 인덱스
     */
	protected void renderHeadersWithNewSheet(Sheet sheet, int rowIndex, int columnStartIndex) {
        // 시트 생성 후 행 추가 전에 auto size을 위한 tracking 설정
        ((SXSSFSheet) sheet).trackAllColumnsForAutoSizing();

        // Auto Size 설정해도 컬럼 너비가 정확하지 않은 경우가 있어 추가 너비 세팅
        ((SXSSFSheet) sheet).setArbitraryExtraWidth(COLUMN_WIDTH_PADDING);

        ExcelHeader excelHeader = resource.getExcelHeader();

        // 헤더 전체 높이
        int headerHeight = excelHeader.getHeaderHeight();
        // 헤더 전체 높이만큼 Row 생성
        for(int rowDepth = 0; rowDepth < headerHeight; rowDepth++){
            if(sheet.getLastRowNum() < rowIndex + rowDepth){
                sheet.createRow(rowIndex + rowDepth);
            }
        }

        // 헤더 Cell 생성
        for(String fieldPath : resource.getFieldPaths()){
            ExcelHeaderCell excelHeaderCell = excelHeader.getExcelHeaderCell(fieldPath);
            excelHeaderCell.adjustCellRange(rowIndex, columnStartIndex);

            Row row = sheet.getRow(excelHeaderCell.getFirstRow());
            Cell cell = row.createCell(excelHeaderCell.getFirstColumn());

            cell.setCellValue(excelHeaderCell.getHeaderName());
            cell.setCellStyle(resource.getCellStyle(fieldPath, ExcelRenderLocation.HEADER));

            // 하나 이상 셀이 있는 경우 셀 병합
            if(excelHeaderCell.isMoreThanOneCell()){
                sheet.addMergedRegion(new CellRangeAddress(excelHeaderCell.getFirstRow(), excelHeaderCell.getLastRow(),
                        excelHeaderCell.getFirstColumn(), excelHeaderCell.getLastColumn()));
            }
        }

        // 병합한 셀 테두리 Border THIN 설정
        setBordersToMergedCells(sheet);
	}

    /**
     * 주어진 데이터를 기반으로 Excel 시트의 본문을 렌더링합니다.
     * 각각의 필드 경로에 해당하는 데이터를 순회하며, 셀에 값을 채우고 스타일을 적용합니다.
     *
     * @param data 본문에 렌더링할 데이터 객체
     * @param rowIndex 렌더링이 시작될 행 인덱스
     * @param columnStartIndex 렌더링이 시작될 열 인덱스
     */
	protected void renderBody(Object data, int rowIndex, int columnStartIndex) {
        Row row = sheet.createRow(rowIndex);
        int columnIndex = columnStartIndex;

        for(String fieldPath : resource.getFieldPaths()){
            if(!resource.getLeafFieldPaths().contains(fieldPath)){
               continue;
            }
            Cell cell = row.createCell(columnIndex++);
            try{
                Object cellValue = getDataValueByFieldPath(fieldPath, data);

                cell.setCellStyle(resource.getCellStyle(fieldPath, ExcelRenderLocation.BODY));
                renderCellValue(cell, cellValue);
            }catch (Exception e){
                throw new ExcelInternalException(e.getMessage(), e);
            }
        }
	}

    /**
     * 주어진 셀(Cell)에 값을 렌더링합니다. 값은 다양한 타입(Number, List, 기타 객체 등)에 따라
     * 적절한 형태로 변환된 후 셀에 설정됩니다.
     *
     * @param cell 값을 설정할 대상 셀(Cell) 객체
     * @param cellValue 셀에 설정할 값, Number, List 또는 기타 객체를 포함할 수 있음
     */
	private void renderCellValue(Cell cell, Object cellValue) {
		// Number 형식
        if (cellValue instanceof Number numberValue) {
            cell.setCellValue(numberValue.doubleValue());
			return;
		}

        // List 형식 지정
        if (cellValue instanceof List<?> listValue) {
            String formattedList = formatListValue(listValue);
            cell.setCellValue(formattedList);
            return;
        }

		cell.setCellValue(cellValue == null ? "" : cellValue.toString());
	}

    /**
     * 주어진 OutputStream에 엑셀 데이터를 쓰고, 관련 리소스를 정리합니다.
     *
     * @param stream 데이터를 작성할 OutputStream 객체
     * @throws IOException 출력 과정에서 입출력 오류가 발생할 경우
     */
	public void write(OutputStream stream) throws IOException {
        wb.write(stream);
		wb.close();
		stream.close();
	}

    /**
     * List 값을 설정된 구분자로 포맷팅합니다.
     * @param listValue 포맷팅할 List
     * @return 구분자로 연결된 문자열
     */
    private String formatListValue(List<?> listValue) {
        if (listValue == null || listValue.isEmpty()) {
            return "";
        }

        return listValue.stream()
                .map(item -> item == null ? "" : item.toString())
                .collect(Collectors.joining(listSeparator));
    }

    /**
     * 현재 시트의 열 너비를 자동으로 조정합니다.
     * 현재 시트에서 첫 번째 행의 셀을 기준으로 열 너비를 자동 조정합니다.
     * 자동 조정 후 추가 여유 공간(COLUMN_WIDTH_PADDING)을 더하여 설정합니다.
     *   - autoSize만으로는 열 너비가 정확하게 조정되지 않아 추가 여유 공간 설정
     */
    protected void autoSizeCurrentSheet() {
        if (sheet != null && sheet.getPhysicalNumberOfRows() > 0) {
            Row row = sheet.getRow(sheet.getFirstRowNum());
            if (row != null) {
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    int columnIndex = cell.getColumnIndex();
                    sheet.autoSizeColumn(columnIndex, true);
                }
            }
        }
    }

    /**
     * 주어진 데이터 객체에서 필드 경로(fieldPath)를 따라 해당 필드의 값을 반환합니다.
     * 필드 경로는 쉼표(",")로 구분되며, 각 경로는 데이터 객체의 필드 이름을 나타냅니다.
     *
     * @param fieldPath 쉼표(",")로 구분된 필드 경로 문자열
     * @param data 값을 추출할 데이터 객체
     * @return 필드 경로에 해당하는 값
     * @throws Exception 필드 접근 중 발생할 수 있는 예외
     */
    private Object getDataValueByFieldPath(String fieldPath, Object data) throws Exception{
        Queue<String> fieldNameQueue = new LinkedList<>(Arrays.asList(fieldPath.split(",")));
        Object result = data;
        Field field = null;

        while(!fieldNameQueue.isEmpty()){
            String fieldName = fieldNameQueue.poll();
            field = getField(data.getClass(), fieldName);
            field.setAccessible(true);
            result = field.get(result);
        }

        return result;
    }

    /**
     * 주어진 시트에서 병합된 모든 셀 영역의 테두리를 설정합니다.
     * 테두리는 상단, 좌측, 우측, 하단 모두 얇은(BorderStyle.THIN) 스타일로 지정됩니다.
     *
     * @param sheet 테두리를 설정할 병합된 셀이 포함된 시트
     */
    private void setBordersToMergedCells(Sheet sheet) {
        List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
        for (CellRangeAddress rangeAddress : mergedRegions) {
            RegionUtil.setBorderTop(BorderStyle.THIN, rangeAddress, sheet);
            RegionUtil.setBorderLeft(BorderStyle.THIN, rangeAddress, sheet);
            RegionUtil.setBorderRight(BorderStyle.THIN, rangeAddress, sheet);
            RegionUtil.setBorderBottom(BorderStyle.THIN, rangeAddress, sheet);
        }
    }
}
