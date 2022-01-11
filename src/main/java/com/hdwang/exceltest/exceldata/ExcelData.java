package com.hdwang.exceltest.exceldata;

import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.cell.CellLocation;

import java.util.List;
import java.util.Map;

/**
 * 表格数据对象
 *
 * @param <T>
 */
public class ExcelData<T> {

    /**
     * 两层list表示的行列数据
     */
    private List<List<CellData>> rowDataList;

    /**
     * Map表示的单元格数据
     */
    private Map<String, CellData> cellDataMap;

    /**
     * 普通javabean表示的数据
     */
    private List<T> beanList;


    public List<List<CellData>> getRowDataList() {
        return rowDataList;
    }

    public void setRowDataList(List<List<CellData>> rowDataList) {
        this.rowDataList = rowDataList;
    }

    public Map<String, CellData> getCellDataMap() {
        return cellDataMap;
    }

    public void setCellDataMap(Map<String, CellData> cellDataMap) {
        this.cellDataMap = cellDataMap;
    }

    public List<T> getBeanList() {
        return beanList;
    }

    public void setBeanList(List<T> beanList) {
        this.beanList = beanList;
    }

    /**
     * 获取指定单元格数据
     *
     * @param rowIndex 行号（从0开始）
     * @param colIndex 列号（从0开始）
     * @return 单元格数据
     */
    public CellData getCellData(int rowIndex, int colIndex) {
        String key = rowIndex + "_" + colIndex;
        return cellDataMap.get(key);
    }

    /**
     * 获取指定单元格数据
     *
     * @param rowIndex 行号（从0开始）
     * @param colName  列名称（从A开始）
     * @return 单元格数据
     */
    public CellData getCellData(int rowIndex, String colName) {
        int colIndex = ExcelUtil.colNameToIndex(colName + "0");
        return getCellData(rowIndex, colIndex);
    }

    /**
     * 获取指定单元格数据
     *
     * @param locationRef 单元格位置（例：A1）
     * @return 单元格数据
     */
    public CellData getCellData(String locationRef) {
        CellLocation cellLocation = ExcelUtil.toLocation(locationRef);
        return getCellData(cellLocation.getY(), cellLocation.getX());
    }

}
