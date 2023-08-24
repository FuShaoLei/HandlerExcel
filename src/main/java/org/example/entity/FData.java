package org.example.entity;

import java.util.List;

public class FData {
    private String code;

    /**
     * 字段名
     */
    private List<String> filedList;

    private List<RowAndColumn> rowAndColumnList;

    public FData() {
    }

    public FData(String code) {
        this.code = code;
    }

    public static FData statusFData(){
        return new FData("status");
    }

    public static FData testDataFData(){
        return new FData("test_data");
    }

    public String getCode() {
        return code;
    }

    public void setCode(String code) {
        this.code = code;
    }

    public List<String> getFiledList() {
        return filedList;
    }

    public void setFiledList(List<String> filedList) {
        this.filedList = filedList;
    }

    public List<RowAndColumn> getRowAndColumnList() {
        return rowAndColumnList;
    }

    public void setRowAndColumnList(List<RowAndColumn> rowAndColumnList) {
        this.rowAndColumnList = rowAndColumnList;
    }
}
