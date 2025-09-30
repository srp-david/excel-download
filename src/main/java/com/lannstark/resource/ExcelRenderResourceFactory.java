package com.lannstark.resource;

import com.lannstark.DefaultBodyStyle;
import com.lannstark.DefaultHeaderStyle;
import com.lannstark.ExcelColumn;
import com.lannstark.ExcelColumnStyle;
import com.lannstark.exception.InvalidExcelCellStyleException;
import com.lannstark.exception.NoExcelColumnAnnotationsException;
import com.lannstark.resource.collection.PreCalculatedCellStyleMap;
import com.lannstark.style.ExcelCellStyle;
import com.lannstark.style.NoExcelCellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.*;

import static com.lannstark.utils.SuperClassReflectionUtils.getAnnotation;

public final class ExcelRenderResourceFactory {

	public static ExcelRenderResource prepareRenderResource(Class<?> type, Workbook wb,
															DataFormatDecider dataFormatDecider) {
		PreCalculatedCellStyleMap styleMap = new PreCalculatedCellStyleMap(dataFormatDecider);
        ExcelHeader excelHeader = new ExcelHeader();
        List<String> fieldPaths = new ArrayList<>();
        List<String> leafFieldPaths = new ArrayList<>();

        // 재귀를 활용하여 전체 헤더 높이 계산하여 재활용
        // max 값이 엑셀 파일의 헤더 높이 결정에 기준이 됨
        int totalHeaderHeight = getHeightOfHeader(type);
        excelHeader.setHeaderHeight(totalHeaderHeight);

		ExcelColumnStyle classDefinedHeaderStyle = getHeaderExcelColumnStyle(type);
		ExcelColumnStyle classDefinedBodyStyle = getBodyExcelColumnStyle(type);

        Queue<FieldPathInfo> fieldPathInfoQueue = new LinkedList<>();

        // 엑셀 대상 DTO에서 ExcelColumn 어노테이션이 있는 필드를 추가
        // BFS 너비 우선 탐색 활용
        List<FieldPathInfo> fieldPathInfos = Arrays.stream(type.getDeclaredFields())
                .filter(field -> field.isAnnotationPresent(ExcelColumn.class))
                .map(field -> getFieldPathInfo("", field)).toList();

        fieldPathInfoQueue.addAll(fieldPathInfos);

        int currRow = 0;
        int currDepth = 1;

        while(!fieldPathInfoQueue.isEmpty()){
            int currCol = 0;

            int mainFieldSize = fieldPathInfoQueue.size();

            for(int i = 0; i < mainFieldSize; i++){
                FieldPathInfo fieldInfo = fieldPathInfoQueue.poll();
                String currFieldPath = fieldInfo.getFieldPath();
                Field currField = fieldInfo.getField();

                // 자식 노드에 추가적으로 ExcelColumn 어노테이션이 붙은 경우
                // 추가 탐색이 필요한 경우
                List<FieldPathInfo> childFieldInfos = Arrays.stream(currField.getType().getDeclaredFields())
                        .filter(child -> child.isAnnotationPresent(ExcelColumn.class))
                        .map(child -> getFieldPathInfo(currFieldPath, child)).toList();

                // 자식 노드 탐색 결과를 Queue에 다시 추가
                fieldPathInfoQueue.addAll(childFieldInfos);

                // FieldPath 목록에 추가
                fieldPaths.add(currFieldPath);

                // 추가 탐색할 게 없는 경우에 추가
                if(childFieldInfos.isEmpty()){
                    leafFieldPaths.add(currFieldPath);
                }

                // ExcelColumn 어노테이션
                ExcelColumn annotation = currField.getAnnotation(ExcelColumn.class);

                // styleMap에 header 정보 추가
                styleMap.put(
                        String.class,
                        ExcelCellKey.of(currFieldPath, ExcelRenderLocation.HEADER),
                        getCellStyle(decideAppliedStyleAnnotation(classDefinedHeaderStyle, annotation.headerStyle())),
                        wb
                );

                // styleMap에 body 정보 추가
                Class<?> currFieldType = currField.getType();
                styleMap.put(
                        currFieldType,
                        ExcelCellKey.of(currFieldPath, ExcelRenderLocation.BODY),
                        getCellStyle(decideAppliedStyleAnnotation(classDefinedBodyStyle, annotation.bodyStyle())),
                        wb
                );

                // 현재 기준으로 자식 노드 갯수
                int childHeaderCount = childFieldInfos.size();

                // childHeaderSize가 0인 경우는 현재 노드가 마지막 노드인 경우임 - 리프 노드
                //  - 위 경우에는 수직으로 병합, 수평 병합은 하지 않고 기본 1 넓이 세팅
                // childHeaderSize > 0 인 경우는 현재 노드가 중간 노드인 경우임 - 자식 노드가 존재
                //  - 위 경우에는 수직 병합 하지 않고 기본 높이 1 세팅, 수평 병합은 자식 노드만큼 진행
                int rowHeight = childHeaderCount == 0 ? totalHeaderHeight - currDepth + 1 : 1;
                int colSpan = childHeaderCount == 0 ? 1 : childHeaderCount;

                // lastRow, lastColumn에서 -1 처리하는 이유는 poi에서 셀 병합 사용 시 index 기준으로 하기 때문에
                ExcelHeaderCell excelHeaderCell = new ExcelHeaderCell(
                        annotation.headerName(),
                        currRow,
                        currRow + rowHeight - 1,
                        currCol,
                        currCol + colSpan - 1
                );

                excelHeader.put(currFieldPath, excelHeaderCell);

                // 현재 Column 위치 변경
                // - 위 쪽에서 표시한 Column 다음에 다음 Column이 표시되어야 하기 때문에
                currCol += colSpan;
            }

            // Row의 인덱스와 Column의 Depth 변경
            currRow++;
            currDepth++;
        }

        if(styleMap.isEmpty()){
            throw new NoExcelColumnAnnotationsException(String.format("Class %s has not @ExcelColumn at all", type));
        }

        return new ExcelRenderResource(styleMap, excelHeader, fieldPaths, leafFieldPaths);
    }

	private static ExcelColumnStyle getHeaderExcelColumnStyle(Class<?> clazz) {
		Annotation annotation = getAnnotation(clazz, DefaultHeaderStyle.class);
		if (annotation == null) {
			return null;
		}
		return ((DefaultHeaderStyle) annotation).style();
	}

	private static ExcelColumnStyle getBodyExcelColumnStyle(Class<?> clazz) {
		Annotation annotation = getAnnotation(clazz, DefaultBodyStyle.class);
		if (annotation == null) {
			return null;
		}
		return ((DefaultBodyStyle) annotation).style();
	}

	private static ExcelColumnStyle decideAppliedStyleAnnotation(ExcelColumnStyle classAnnotation,
																 ExcelColumnStyle fieldAnnotation) {
		if (fieldAnnotation.excelCellStyleClass().equals(NoExcelCellStyle.class) && classAnnotation != null) {
			return classAnnotation;
		}
		return fieldAnnotation;
	}

	private static ExcelCellStyle getCellStyle(ExcelColumnStyle excelColumnStyle) {
		Class<? extends ExcelCellStyle> excelCellStyleClass = excelColumnStyle.excelCellStyleClass();
		// 1. Case of Enum
		if (excelCellStyleClass.isEnum()) {
			String enumName = excelColumnStyle.enumName();
			return findExcelCellStyle(excelCellStyleClass, enumName);
		}

		// 2. Case of Class
		try {
			return excelCellStyleClass.newInstance();
		} catch (InstantiationException | IllegalAccessException e) {
			throw new InvalidExcelCellStyleException(e.getMessage(), e);
		}
	}

	@SuppressWarnings("unchecked")
	private static ExcelCellStyle findExcelCellStyle(Class<?> excelCellStyles, String enumName) {
		try {
			return (ExcelCellStyle) Enum.valueOf((Class<Enum>) excelCellStyles, enumName);
		} catch (NullPointerException e) {
			throw new InvalidExcelCellStyleException("enumName must not be null", e);
		} catch (IllegalArgumentException e) {
			throw new InvalidExcelCellStyleException(
					String.format("Enum %s does not name %s", excelCellStyles.getName(), enumName), e);
		}
	}

    /**
     * 주어진 클래스에서 ExcelColumn 어노테이션이 붙은 필드를 기준으로 클래스의 최대 깊이를 계산합니다.
     * 현재 깊이(currDepth)를 기준으로 재귀적으로 탐색하며 최대 깊이를 반환합니다.
     *
     * @param clazz 깊이를 계산할 기준 클래스
     * @param currDepth 현재 깊이, 탐색 과정에서 증가
     * @return 계산된 클래스의 최대 깊이
     */
    private static int getMaxDepth(Class<?> clazz, int currDepth){
        // ExcelColumn 어노테이션이 붙은 필드 필터링
        List<Field> excelColumnList =
                Arrays.stream(clazz.getDeclaredFields())
                        .filter(field -> field.isAnnotationPresent(ExcelColumn.class))
                        .toList();

        // 어노테이션 붙은 필드 없는 경우
        if(excelColumnList.isEmpty()){
            return currDepth;
        }

        int maxDepth = currDepth;
        for(Field field : excelColumnList){
            maxDepth = Math.max(maxDepth, getMaxDepth(field.getType(), currDepth + 1));
        }
        return maxDepth;
    }

    /**
     * 주어진 클래스의 헤더 높이를 반환합니다.
     * 클래스의 계층 구조에서 ExcelColumn 어노테이션이 붙은 필드를 기준으로
     * 최대 깊이를 계산하여 헤더 높이를 식별합니다.
     *
     * @param clazz 헤더 높이를 계산할 대상 클래스
     * @return 헤더의 높이 값 (최대 깊이 값)
     */
    private static int getHeightOfHeader(Class<?> clazz){
        return getMaxDepth(clazz, 0);
    }

    /**
     * 주어진 필드 경로와 필드를 기반으로 FieldPathInfo 객체를 생성하여 반환합니다.
     * 필드 경로가 비어있지 않으면 필드 경로와 필드 이름을 조합하여 설정하고,
     * 비어있으면 필드 이름만 필드 경로로 설정합니다.
     *
     * @param fieldPath 필드의 경로를 나타내는 문자열
     * @param field Field 객체로, 경로에 포함될 특정 필드를 나타냅니다
     * @return 필드 경로 및 관련 필드 정보를 포함하는 FieldPathInfo 객체
     */
    private static FieldPathInfo getFieldPathInfo(String fieldPath, Field field) {
        if (!fieldPath.isEmpty()) {
            return new FieldPathInfo(String.format("%s,%s", fieldPath, field.getName()), field);
        } else {
            return new FieldPathInfo(field.getName(), field);
        }
    }
}
