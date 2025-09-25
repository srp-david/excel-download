package com.lannstark.header;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.lannstark.ExcelColumn;
import com.lannstark.dto.EmployeeMainDto;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.assertj.core.api.Assertions;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.lang.reflect.Field;
import java.util.*;

public class HeaderTest {
    private final ObjectMapper objectMapper = new ObjectMapper();

    @Test
    @DisplayName("헤더 높이 계산")
    public void getMaxHeight(){
        int maxHeaderHeight = getHeightOfHeader(EmployeeMainDto.class);

        Assertions.assertThat(maxHeaderHeight).isEqualTo(2);
    }

    @Test
    @DisplayName("헤더 생성 테스트")
    public void headerCreationTest() throws Exception {
        int maxHeaderHeight = getHeightOfHeader(EmployeeMainDto.class);
        List<HeaderCellInfo> headerCellInfos = createHeader(EmployeeMainDto.class, 1, 1, maxHeaderHeight);
        System.out.println("header: " + objectMapper.writerWithDefaultPrettyPrinter().writeValueAsString(headerCellInfos));
    }

    private List<HeaderCellInfo> createHeader(Class<?> clazz, int firstRow, int firstColumn, int maxHeight){
        List<HeaderCellInfo> headerCellInfos = new ArrayList<>();

        Queue<Field> fieldQueue = new LinkedList<>();
        fieldQueue.addAll(Arrays.asList(clazz.getDeclaredFields()));

        int currRow = firstRow;
        int currDepth = 1;

        while(!fieldQueue.isEmpty()){
            int currCol = firstColumn;

            int queueSize = fieldQueue.size();
            for(int i = 0; i < queueSize; i++){
                Field field = fieldQueue.poll();

                int childHeaderSize = 0;
                for (Field child : field.getType().getDeclaredFields()) {
                    // 자식 요소의 개수 구하기
                    if (child.isAnnotationPresent(ExcelColumn.class)) {
                        childHeaderSize++;
                    }

                    // 추가적으로 탐색해야 하는 자식 요소는 Queue에 넣기
                    if (child.isAnnotationPresent(ExcelColumn.class)) {
                        fieldQueue.add(child);
                    }
                }

                int rowHeight = childHeaderSize == 0 ? maxHeight - currDepth + 1 : 1;
                int colSpan = childHeaderSize == 0 ? 1 : childHeaderSize;

                HeaderCellInfo headerCellInfo = new HeaderCellInfo(
                        field.getName(),
                        currRow,
                        currRow + rowHeight - 1,
                        currCol,
                        currCol + colSpan - 1
                );

                headerCellInfos.add(headerCellInfo);

                currCol += colSpan;
            }

            currRow++;
            currDepth++;
        }

        return headerCellInfos;
    }

    private int getHeightOfHeader(Class<?> clazz){
        return getDepth(clazz, 0);
    }

    private int getDepth(Class<?> clazz, int currDepth){
        List<Field> fieldList = Arrays.stream(clazz.getDeclaredFields())
                .filter(field -> field.isAnnotationPresent(ExcelColumn.class))
                .toList();

        if(fieldList.isEmpty()){
            return currDepth;
        }

        int maxDepth = currDepth;
        for(Field field : fieldList){
            maxDepth = Math.max(maxDepth, getDepth(field.getType(), currDepth+1));
        }

        return maxDepth;
    }

    @Getter
    @Setter
    @NoArgsConstructor
    @AllArgsConstructor
    private class HeaderCellInfo {
       private String headerName;
       private int firstRow;
       private int lastRow;
       private int firstColumn;
       private int lastColumn;
    }
}