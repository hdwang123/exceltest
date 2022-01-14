package com.hdwang.exceltest.exceldata;

import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.cell.CellHandler;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.Cell;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * 表格数据工具类
 */
public class ExcelDataReader {

    /**
     * 读取表格数据
     *
     * @param templateFile   文件
     * @param startRowIndex  起始行号（从0开始）
     * @param endRowIndex    结束行号（从0开始,包括此列）
     * @param startCellIndex 起始列号（从0开始）
     * @param endCellIndex   结束列号（从0开始,包括此列）
     * @return 表格数据
     */
    public static ExcelData readExcelData(File templateFile, int startRowIndex, int endRowIndex, int startCellIndex, int endCellIndex) {
        ExcelData excelData = new ExcelData();
        ExcelReader excelReader = ExcelUtil.getReader(templateFile, 0);
        List<List<CellData>> rowDataList = new ArrayList<>();
        AtomicInteger rowIndex = new AtomicInteger(-1);
        //读取表格数据
        excelReader.read(startRowIndex, endRowIndex, new CellHandler() {
            @Override
            public void handle(Cell cell, Object value) {
                if (cell == null) {
                    //无单元格跳过
                    return;
                }
                if (cell.getColumnIndex() < startCellIndex || cell.getColumnIndex() > endCellIndex) {
                    //列号不在范围内跳过
                    return;
                }

                //新行的数据
                if (cell.getRowIndex() != rowIndex.get()) {
                    rowDataList.add(new ArrayList<>());
                }
                rowIndex.set(cell.getRowIndex());
                //取出新行数据对象存储单元格数据
                List<CellData> cellDataList = rowDataList.get(rowDataList.size() - 1);
                CellData cellData = new CellData();
                cellData.setRowIndex(cell.getRowIndex());
                cellData.setCellIndex(cell.getColumnIndex());
                cellData.setValue(value);
                cellDataList.add(cellData);
            }
        });
        excelData.setRowDataList(rowDataList);
        //转换为Map结构
        excelData.setCellDataMap(convertExcelDataToMap(rowDataList));
        return excelData;
    }

    /**
     * 读取表格数据
     *
     * @param templateFile  文件
     * @param startRowIndex 起始行号（从0开始）
     * @param endRowIndex   结束行号（从0开始,包括此列）
     * @param startCellName 起始列名（从A开始）
     * @param endCellName   结束列名（从A开始,包括此列）
     * @return 表格数据
     */
    public static ExcelData readExcelData(File templateFile, int startRowIndex, int endRowIndex, String startCellName, String endCellName) {
        int startCellIndex = ExcelUtil.colNameToIndex(startCellName + "0");
        int endCellIndex = ExcelUtil.colNameToIndex(endCellName + "0");
        return readExcelData(templateFile, startRowIndex, endRowIndex, startCellIndex, endCellIndex);
    }

    /**
     * 转换表格数据为Map结构，用于数据快速查找
     *
     * @param rowDataList 表格数据
     * @return Map结构表示的表格数据
     */
    private static Map<String, CellData> convertExcelDataToMap(List<List<CellData>> rowDataList) {
        if (CollectionUtils.isEmpty(rowDataList)) {
            return new HashMap<>();
        }
        Map<String, CellData> cellDataMap = new HashMap<>();
        for (List<CellData> rowData : rowDataList) {
            for (CellData cellData : rowData) {
                String key = cellData.getRowIndex() + "_" + cellData.getCellIndex();
                cellDataMap.put(key, cellData);
            }
        }
        return cellDataMap;
    }


}
