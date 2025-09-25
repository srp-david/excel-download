package com.lannstark.body;

import com.lannstark.dto.DeptInfo;
import com.lannstark.dto.EmployeeInfo;
import com.lannstark.dto.EmployeeMainDto;
import org.assertj.core.api.Assertions;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.lang.reflect.Field;
import java.util.Arrays;
import java.util.LinkedList;
import java.util.Queue;

import static com.lannstark.utils.SuperClassReflectionUtils.getField;

public class BodyTest {

    @Test
    @DisplayName("Body 생성 테스트")
    public void bodyCreationTest() throws Exception{

        EmployeeMainDto mainDto = new EmployeeMainDto(
                new EmployeeInfo(
                        "David",
                        29
                ),
                new DeptInfo(
                        "전산실",
                        "DEPT-0001",
                        "(주)에스알피인포텍"
                )
        );

        Object dtoValue = getDtoValue("employInfo,name", mainDto);

        Assertions.assertThat(dtoValue).isEqualTo("David");
    }

    private static Object getDtoValue(String fieldPath, Object mainDto) throws Exception{
        Queue<String> fieldPathQueue = new LinkedList<>(Arrays.asList(fieldPath.split(",")));
        Object result = mainDto;
        Field field = null;
        while(!fieldPathQueue.isEmpty()){
            String fieldPathItem = fieldPathQueue.poll();
            field = getField(result.getClass(), fieldPathItem);
            field.setAccessible(true);
            result = field.get(result);
        }
        return result;
    }
}