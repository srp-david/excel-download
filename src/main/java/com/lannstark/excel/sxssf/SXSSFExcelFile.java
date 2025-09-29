package com.lannstark.excel.sxssf;

import com.lannstark.excel.ExcelFile;
import com.lannstark.exception.ExcelInternalException;
import com.lannstark.resource.*;
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

import static com.lannstark.utils.SuperClassReflectionUtils.getField;

/**
 * SXSSFExcelFile 클래스는 Apache POI의 SXSSF(SXSSFWorkbook)를 이용하여 Excel 파일을 생성하고 데이터의 렌더링을 지원하는 추상 클래스입니다.
 *  - 템플릿 메서드 패턴 기반
 * @param <T> 렌더링할 데이터 타입
 */
public abstract class SXSSFExcelFile<T> implements ExcelFile<T> {

	protected static final SpreadsheetVersion supplyExcelVersion = SpreadsheetVersion.EXCEL2007;
	private static final int COLUMN_WIDTH_PADDING = 2500;

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
	 * @param data List Data to render an Excel file. data should have at least one @ExcelColumn on fields
	 * @param type Class type to be rendered
	 */
	public SXSSFExcelFile(List<T> data, Class<T> type) {
		this(data, type, new DefaultDataFormatDecider());
	}

	/**
	 * SXSSFExcelFile
	 * @param data List Data to render an Excel file. data should have at least one @ExcelColumn on fields
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
     * 제공된 데이터를 바탕으로 Excel 파일을 렌더링합니다.
     *
     * @param data Excel 파일에 렌더링할 데이터 리스트
     */
	protected abstract void renderExcel(List<T> data);

    /**
     * 현재 시트의 열 너비를 자동으로 조정합니다.
     *
     * 현재 시트에서 첫 번째 행의 셀을 기준으로 열 너비를 자동 조정합니다.
     * 자동 조정 후 추가 여유 공간(COLUMN_WIDTH_PADDING)을 더하여 설정합니다.
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
                    int currentColumnWidth = sheet.getColumnWidth(columnIndex);
                    sheet.setColumnWidth(columnIndex, (currentColumnWidth + COLUMN_WIDTH_PADDING));
                }
            }
        }
    }

    /**
     * 새 시트에 헤더를 생성하고 렌더링합니다.
     *
     * 주어진 시트 객체를 기반으로 헤더를 생성하며, 해당 헤더는 ExcelHeader와 매핑된 필드 경로를 사용하여 생성됩니다.
     * 헤더의 셀 병합 및 스타일 지정 작업도 이 메서드에서 수행됩니다.
     *
     * @param sheet 헤더를 생성할 대상 시트
     * @param rowIndex 시작 행 인덱스
     * @param columnStartIndex 시작 열 인덱스
     */
	protected void renderHeadersWithNewSheet(Sheet sheet, int rowIndex, int columnStartIndex) {
        ((SXSSFSheet) sheet).trackAllColumnsForAutoSizing();

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

	private void renderCellValue(Cell cell, Object cellValue) {
		if (cellValue instanceof Number) {
			Number numberValue = (Number) cellValue;
			cell.setCellValue(numberValue.doubleValue());
			return;
		}
		cell.setCellValue(cellValue == null ? "" : cellValue.toString());
	}

	public void write(OutputStream stream) throws IOException {
        wb.write(stream);
		wb.close();
		stream.close();
	}

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
