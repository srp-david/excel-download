package com.lannstark.excel.sxssf.multiplesheet;

import com.lannstark.excel.sxssf.SXSSFExcelFile;
import com.lannstark.resource.DataFormatDecider;
import org.apache.commons.compress.archivers.zip.Zip64Mode;
import org.apache.commons.lang3.StringUtils;

import java.util.List;

/**
 * 이 클래스는 Apache POI 라이브러리를 사용하여 Excel 파일을 생성하며,
 * 다수의 시트(Sheet)를 지원하는 기능을 제공합니다.
 * 제네릭(Generic) 타입을 사용하여 데이터 모델을 정의하고,
 * 주어진 데이터를 Excel 파일에 렌더링합니다.
 *
 * @param <T> Excel에 렌더링할 데이터의 제네릭 타입
 */
public class MultiSheetExcelFile<T> extends SXSSFExcelFile<T> {

	private static final int maxRowCanBeRendered = supplyExcelVersion.getMaxRows() - 1;
	private static final int ROW_START_INDEX = 0;
	private static final int COLUMN_START_INDEX = 0;
	private int currentRowIndex = ROW_START_INDEX;

    private String baseSheetName;
    private int sheetIndex;

	public MultiSheetExcelFile(Class<T> type) {
		super(type);
		wb.setZip64Mode(Zip64Mode.Always);
		initializeFields();
	}

	/*
	 * If you use SXSSF with hug data, you need to set zip mode
	 * see http://apache-poi.1045710.n5.nabble.com/Bug-62872-New-Writing-large-files-with-800k-rows-gives-java-io-IOException-This-archive-contains-unc-td5732006.html
	 */
	public MultiSheetExcelFile(List<T> data, Class<T> type) {
		super(data, type);
		wb.setZip64Mode(Zip64Mode.Always);
		initializeFields();
	}

	public MultiSheetExcelFile(List<T> data, Class<T> type, DataFormatDecider dataFormatDecider) {
		super(data, type, dataFormatDecider);
		wb.setZip64Mode(Zip64Mode.Always);
		initializeFields();
	}

    /**
     * 필드를 초기화합니다.
     */
    private void initializeFields() {
        this.baseSheetName = "Sheet";
        this.sheetIndex = 1;
    }

    /**
     * 데이터를 기반으로 Excel 파일에 내용을 렌더링합니다.
     * 데이터가 비어 있는 경우, 새로운 시트를 생성하고 헤더만 추가합니다. 그렇지 않으면 데이터 행을 추가합니다.
     *
     * @param data Excel 파일에 렌더링할 데이터 목록
     */
	@Override
	protected void renderExcel(List<T> data) {
		// 1. Create Header and return if data is empty
		if (data.isEmpty()) {
			createNewSheetWithHeader();
			autoSizeCurrentSheet();
			return ;
		}

		// 2. Render body
		createNewSheetWithHeader();
		addRows(data);
	}

    /**
     * 주어진 데이터를 기반으로 Excel 파일에 행을 추가합니다.
     * 데이터의 끝에 도달하거나 현재 시트의 최대 행 수를 초과하면
     * 새로운 시트를 생성하고 헤더를 추가한 후 이어서 행을 렌더링합니다.
     *
     * @param data Excel 파일에 추가할 데이터 목록
     */
	@Override
	public void addRows(List<T> data) {
		for (Object renderedData : data) {
			renderBody(renderedData, currentRowIndex++, COLUMN_START_INDEX);

			if (currentRowIndex == maxRowCanBeRendered) {
				autoSizeCurrentSheet();
				currentRowIndex = 1;
				createNewSheetWithHeader();
			}
		}

		// 마지막 시트에 대한 auto sizing
		autoSizeCurrentSheet();
	}

    /**
     * 새 시트를 생성하고 헤더를 렌더링합니다.
     * 현재 워크북에 새로운 시트를 추가하며, 시트 이름은 기본 시트 이름과
     * 시트 인덱스를 조합하여 설정됩니다. 이후, 새롭게 생성된 시트의 지정된
     * 행 시작 인덱스와 열 시작 인덱스 위치에 헤더를 렌더링합니다.
     * 또한, 현재 행 인덱스를 초기화하거나 다음 데이터 추가 작업을
     * 준비하기 위해 1 증가시킵니다.
     * 이 메서드는 데이터가 없는 경우 헤더만 생성하고 렌더링하거나,
     * 현재 시트의 최대 행 제한을 초과한 경우 새 시트를 생성하며 사용됩니다.
     */
    private void createNewSheetWithHeader() {
		sheet = wb.createSheet(baseSheetName + sheetIndex++);

		renderHeadersWithNewSheet(sheet, ROW_START_INDEX, COLUMN_START_INDEX);
		currentRowIndex++;
	}

    /**
     * 기본 시트 이름을 설정합니다.
     * 전달된 값이 비어있지 않은 경우에만 기본 시트 이름을 업데이트합니다.
     *
     * @param baseSheetName 기본 시트 이름으로 설정할 문자열
     */
    public void setSheetName(String baseSheetName) {
        if(StringUtils.isNotEmpty(baseSheetName)){
            this.baseSheetName = baseSheetName;
        }
    }
}
