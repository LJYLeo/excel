package com.codeshare.excel;

import java.util.List;
import java.util.Map;

public class Excel {

    /**
     * 行数，表格行号减1
     */
    private int startRow;

    /**
     * 行数，表格行号减1
     */
    private int endRow;

    /**
     * 要映射的单元格列数，新老模板一一对应
     */
    private List<Integer> cell;

    /**
     * 行列坐标，新老模板一一对应
     */
    private List<Map<String, Integer>> doubleCell;

    public int getStartRow() {
        return startRow;
    }

    public void setStartRow(int startRow) {
        this.startRow = startRow;
    }

    public int getEndRow() {
        return endRow;
    }

    public void setEndRow(int endRow) {
        this.endRow = endRow;
    }

    public List<Integer> getCell() {
        return cell;
    }

    public void setCell(List<Integer> cell) {
        this.cell = cell;
    }

    public List<Map<String, Integer>> getDoubleCell() {
        return doubleCell;
    }

    public void setDoubleCell(List<Map<String, Integer>> doubleCell) {
        this.doubleCell = doubleCell;
    }

    @Override
    public String toString() {
        return "Excel{" +
                "startRow=" + startRow +
                ", endRow=" + endRow +
                ", cell=" + cell +
                ", doubleCell=" + doubleCell +
                '}';
    }

}
