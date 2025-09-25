package com.lannstark.dto;

import com.lannstark.DefaultHeaderStyle;
import com.lannstark.ExcelColumn;
import com.lannstark.ExcelColumnStyle;
import com.lannstark.style.BlackHeaderStyle;
import com.lannstark.style.BlueHeaderStyle;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@DefaultHeaderStyle(style = @ExcelColumnStyle(excelCellStyleClass = BlueHeaderStyle.class))
public class EmployeeMainDto {
    @ExcelColumn(headerName = "직원 정보")
    private EmployeeInfo employInfo;

    @ExcelColumn(headerName = "부서 정보", headerStyle = @ExcelColumnStyle(excelCellStyleClass = BlackHeaderStyle.class))
    private DeptInfo deptInfo;
}