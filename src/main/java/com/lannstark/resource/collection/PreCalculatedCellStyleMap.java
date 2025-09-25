package com.lannstark.resource.collection;

import com.lannstark.resource.DataFormatDecider;
import com.lannstark.resource.ExcelCellKey;
import com.lannstark.style.ExcelCellStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.HashMap;
import java.util.Map;

/**
 * 일급 컬렉션 활용
 * PreCalculatedCellStyleMap 클래스는 Excel의 셀 스타일을 사전에 계산하고 저장하는 역할을 합니다.
 * 주어진 필드 타입과 키 값을 기반으로 CellStyle 객체를 생성하여 맵에 저장하고,
 * 이후 동일한 키를 사용하여 저장된 CellStyle을 빠르게 조회할 수 있습니다.
 */
public class PreCalculatedCellStyleMap {

	private final DataFormatDecider dataFormatDecider;

    private final Map<ExcelCellKey, CellStyle> cellStyleMap = new HashMap<>();

	public PreCalculatedCellStyleMap(DataFormatDecider dataFormatDecider) {
		this.dataFormatDecider = dataFormatDecider;
	}

	public void put(Class<?> fieldType, ExcelCellKey excelCellKey, ExcelCellStyle excelCellStyle, Workbook wb) {
		CellStyle cellStyle = wb.createCellStyle();
		DataFormat dataFormat = wb.createDataFormat();
		cellStyle.setDataFormat(dataFormatDecider.getDataFormat(dataFormat, fieldType));
		excelCellStyle.apply(cellStyle);
		cellStyleMap.put(excelCellKey, cellStyle);
	}

	public CellStyle get(ExcelCellKey excelCellKey) {
		return cellStyleMap.get(excelCellKey);
	}

	public boolean isEmpty() {
		return cellStyleMap.isEmpty();
	}

}
