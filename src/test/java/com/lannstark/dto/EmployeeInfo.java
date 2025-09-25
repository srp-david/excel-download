package com.lannstark.dto;

import com.lannstark.ExcelColumn;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
public class EmployeeInfo {
    @ExcelColumn(headerName = "직원명")
    public String name;
    @ExcelColumn(headerName = "나이")
    public int age;
}