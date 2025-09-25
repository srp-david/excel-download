package com.lannstark.resource;

import java.util.Objects;

/**
 * ExcelCellKey 클래스는 엑셀 셀의 고유 키를 표현하기 위한 불변 객체입니다.
 * 필드 경로와 해당 필드의 렌더링 위치(헤더 또는 바디)에 따라 고유한 키를 생성합니다.
 * 이 키는 엑셀 렌더링 리소스에서 스타일 맵핑 등의 작업에 사용됩니다.
 *
 * 이 클래스는 필드 경로와 렌더링 위치를 결합하여 {@link ExcelRenderResource}에서 셀 스타일을 찾거나 설정하는 데 사용됩니다.
 *
 * 주요 특징:
 * - {@code fieldPath}: 필드를 식별하기 위한 고유 경로
 * - {@code excelRenderLocation}: 렌더링 위치, 헤더 또는 바디를 지정
 * - 불변 객체로 설계되어 데이터를 안전하게 처리할 수 있음
 */
public final class ExcelCellKey {

	private final String fieldPath;
	private final ExcelRenderLocation excelRenderLocation;

	private ExcelCellKey(String fieldPath, ExcelRenderLocation excelRenderLocation) {
		this.fieldPath = fieldPath;
		this.excelRenderLocation = excelRenderLocation;
	}

	public static ExcelCellKey of(String fieldPath, ExcelRenderLocation excelRenderLocation) {
		assert excelRenderLocation != null;
		return new ExcelCellKey(fieldPath, excelRenderLocation);
	}

	@Override
	public boolean equals(Object o) {
		if (this == o) return true;
		if (o == null || getClass() != o.getClass()) return false;
		ExcelCellKey that = (ExcelCellKey) o;
		return Objects.equals(fieldPath, that.fieldPath) &&
				excelRenderLocation == that.excelRenderLocation;
	}

	@Override
	public int hashCode() {
		return Objects.hash(fieldPath, excelRenderLocation);
	}

}
