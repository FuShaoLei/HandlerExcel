package org.example.entity;

public class RowAndColumn {

    private int rowStart;
    private int rowEnd;
    private int specifiedColumn;

    public RowAndColumn(int rowStart, int rowEnd, int specifiedColumn) {
        this.rowStart = rowStart;
        this.rowEnd = rowEnd;
        this.specifiedColumn = specifiedColumn;
    }

    public int getRowStart() {
        return rowStart;
    }

    public void setRowStart(int rowStart) {
        this.rowStart = rowStart;
    }

    public int getRowEnd() {
        return rowEnd;
    }

    public void setRowEnd(int rowEnd) {
        this.rowEnd = rowEnd;
    }

    public int getSpecifiedColumn() {
        return specifiedColumn;
    }

    public void setSpecifiedColumn(int specifiedColumn) {
        this.specifiedColumn = specifiedColumn;
    }
}
