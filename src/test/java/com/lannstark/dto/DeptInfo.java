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
public class DeptInfo {
    @ExcelColumn(headerName = "부서명")
    private String deptName;
    @ExcelColumn(headerName = "부서 코드")
    private String deptCode;
    @ExcelColumn(headerName = "상위 부서")
    private String upDeptName;
}