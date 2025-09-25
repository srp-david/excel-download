package com.lannstark.resource;

import com.lannstark.resource.collection.PreCalculatedCellStyleMap;
import lombok.Getter;
import org.apache.poi.ss.usermodel.CellStyle;

import java.util.List;

/**
 * ExcelRenderResource 클래스는 엑셀 렌더링 과정에서 필요한 리소스를 캡슐화하는 역할을 합니다.
 * 스타일 맵, 헤더 정보, 필드 경로 등의 데이터를 포함하며,
 * 이를 토대로 적절한 셀 스타일과 헤더 데이터를 제공합니다.
 *
 * 이 클래스는 엑셀 파일 생성 과정에서 데이터를 효과적으로 렌더링하기 위한 핵심적인 자원을 제공합니다.
 * 내부적으로 사전에 계산된 셀 스타일 맵과 헤더 정보를 활용하여 빠르고 효율적인 렌더링을 지원합니다.
 *
 * 주요 구성 요소:
 * - {@code styleMap}: 필드 경로와 렌더링 위치를 기준으로 사전에 계산된 셀 스타일이 저장된 맵
 * - {@code excelHeader}: 엑셀 헤더 정보가 포함된 객체
 * - {@code fieldPaths}: 렌더링 대상 필드의 전체 경로 리스트
 * - {@code leafFieldPaths}: 렌더링 대상 필드 중 말단 필드의 경로 리스트
 *
 * 주요 기능:
 * - 특정 필드 경로와 렌더링 위치에 기반하여 해당 셀의 스타일을 반환
 */
@Getter
public class ExcelRenderResource {

	private PreCalculatedCellStyleMap styleMap;
    private ExcelHeader excelHeader;
    private List<String> fieldPaths;
    private List<String> leafFieldPaths;

    public ExcelRenderResource(PreCalculatedCellStyleMap styleMap, ExcelHeader excelHeader, List<String> fieldPaths, List<String> leafFieldPaths) {
        this.styleMap = styleMap;
        this.excelHeader = excelHeader;
        this.fieldPaths = fieldPaths;
        this.leafFieldPaths = leafFieldPaths;
    }

    public CellStyle getCellStyle(String fieldPath, ExcelRenderLocation excelRenderLocation) {
        return styleMap.get(ExcelCellKey.of(fieldPath, excelRenderLocation));
    }

}
