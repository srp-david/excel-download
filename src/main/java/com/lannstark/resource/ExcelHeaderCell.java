package com.lannstark.resource;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
public class ExcelHeaderCell {

    private String headerName;
    private int firstRow;
    private int lastRow;
    private int firstColumn;
    private int lastColumn;

    /**
     * 셀의 행과 열 범위를 조정합니다.
     * 입력받은 시작 인덱스를 기준으로 현재 행과 열 범위를 변경합니다.
     * 셀 병합할 때 필요
     * @param rowStartIndex 셀 범위의 행 시작 인덱스 조정을 위한 값
     * @param columnStartIndex 셀 범위의 열 시작 인덱스 조정을 위한 값
     */
    public void adjustCellRange(int rowStartIndex, int columnStartIndex){
        this.firstRow = firstRow + rowStartIndex;
        this.lastRow = lastRow + rowStartIndex;
        this.firstColumn = firstColumn + columnStartIndex;
        this.lastColumn = lastColumn + columnStartIndex;
    }

    /**
     * 셀 병합 여부 체크
     * true인 경우 병합할 행이나 컬럼이 추가적으로 있는 경우
     * false는 단일 행이나 단일 컬럼인 경우
     * @return 셀이 하나 이상의 행 또는 열을 포함하면 true, 그렇지 않으면 false를 반환
     */
    public boolean isMoreThanOneCell(){
        return lastRow > firstRow || lastColumn > firstColumn;
    }
}