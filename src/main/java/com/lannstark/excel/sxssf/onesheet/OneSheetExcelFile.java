package com.lannstark.excel.sxssf.onesheet;

import com.lannstark.excel.sxssf.SXSSFExcelFile;
import com.lannstark.resource.DataFormatDecider;
import org.apache.commons.lang3.StringUtils;

import java.util.List;

/**
 * OneSheetExcelFile
 * - support Excel Version over 2007
 * - support one sheet rendering
 * - support different DataFormat by Class Type
 * - support Custom CellStyle according to (header or body) and data field
 */
public final class OneSheetExcelFile<T> extends SXSSFExcelFile<T> {

	private static final int ROW_START_INDEX = 0;
	private static final int COLUMN_START_INDEX = 0;
	private int currentRowIndex = ROW_START_INDEX;

    private String sheetName = "Sheet1";

	public OneSheetExcelFile(Class<T> type) {
        super(type);
	}

	public OneSheetExcelFile(List<T> data, Class<T> type) {
		super(data, type);
	}

	public OneSheetExcelFile(List<T> data, Class<T> type, DataFormatDecider dataFormatDecider) {
		super(data, type, dataFormatDecider);
	}

    /**
     * 제공된 데이터가 Excel 파일 구성을 위한 유효한지 검증합니다.
     * 데이터의 크기가 Excel 버전에서 지원하는 최대 행 수를 초과할 경우 예외를 발생시킵니다.
     *
     * @param data 검증할 데이터 목록
     * @throws IllegalArgumentException 데이터의 크기가 Excel 버전의 최대 행 수를 초과할 경우 발생
     */
	@Override
	protected void validateData(List<T> data) {
		int maxRows = supplyExcelVersion.getMaxRows();
		if (data.size() > maxRows) {
			throw new IllegalArgumentException(
					String.format("This concrete ExcelFile does not support over %s rows", maxRows));
		}
	}

    /**
     * 주어진 데이터를 이용해 Excel 파일을 생성하고 렌더링합니다.
     * - 데이터가 비어 있는 경우 헤더만 렌더링됩니다.
     * - 데이터가 존재하는 경우 각 데이터를 반복하며 본문을 렌더링합니다.
     *
     * @param data Excel 파일에 포함될 데이터 목록. 데이터 유형은 제네릭 타입 T를 따릅니다.
     */
	@Override
	public void renderExcel(List<T> data) {
        // 1. Create sheet and renderHeader
		sheet = wb.createSheet(sheetName);
		renderHeadersWithNewSheet(sheet, currentRowIndex++, COLUMN_START_INDEX);

		if (data.isEmpty()) {
			return;
		}

		// 2. Render Body
		for (Object renderedData : data) {
			renderBody(renderedData, currentRowIndex++, COLUMN_START_INDEX);
		}
	}

    /**
     * 데이터가 많은 경우 쪼개서 넣기 위한 용도
     * @param data 데이터
     */
    @Override
    public void addRows(List<T> data) {
        if (currentRowIndex == 0) currentRowIndex = 1;
        for (Object renderedData : data) {
            renderBody(renderedData, currentRowIndex++, COLUMN_START_INDEX);
        }
    }

    /**
     * Excel 시트의 이름을 설정합니다. 주어진 시트 이름이 비어 있지 않은 경우에만 설정되며,
     * 비어 있을 경우 기본값이 유지됩니다.
     *
     * @param sheetName 설정할 시트 이름. null이거나 빈 문자열이 아닌 경우 시트 이름으로 설정됩니다.
     */
    public void setSheetName(String sheetName) {
        if(StringUtils.isNotEmpty(sheetName)){
            this.sheetName = sheetName;
        }
    }
}
