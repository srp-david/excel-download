package com.lannstark.resource;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.HashMap;
import java.util.Map;

@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
public class ExcelHeader {

    // header 높이
    private int headerHeight;
    // fieldPath와 ExcelHeaderCell 맵핑
    private Map<String, ExcelHeaderCell> headerCellMap = new HashMap<>();

    public void put(String fieldPath, ExcelHeaderCell excelHeaderCell){
        this.headerCellMap.put(fieldPath, excelHeaderCell);
    }

    public ExcelHeaderCell getExcelHeaderCell(String fieldPath){
        return this.headerCellMap.get(fieldPath);
    }
}