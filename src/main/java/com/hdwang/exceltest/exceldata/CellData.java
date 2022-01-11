package com.hdwang.exceltest.exceldata;

import cn.hutool.json.JSONUtil;

/**
 * 单元格数据
 */
public class CellData {

    /**
     * 行号
     */
    private int rowIndex;

    /**
     * 列号
     */
    private int cellIndex;

    /**
     * 单元格数值
     */
    private Object value;

    public int getRowIndex() {
        return rowIndex;
    }

    public void setRowIndex(int rowIndex) {
        this.rowIndex = rowIndex;
    }

    public int getCellIndex() {
        return cellIndex;
    }

    public void setCellIndex(int cellIndex) {
        this.cellIndex = cellIndex;
    }

    public Object getValue() {
        return value;
    }

    public void setValue(Object value) {
        this.value = value;
    }

    @Override
    public String toString() {
        return JSONUtil.toJsonStr(this);
    }
}
