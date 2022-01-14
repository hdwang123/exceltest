package com.hdwang.exceltest.exceldata;

import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.cell.CellHandler;
import cn.hutool.poi.excel.cell.CellLocation;
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
     * 读取表格指定行列范围内的数据,默认读取第一个sheet里的内容
     *
     * @param file          文件
     * @param startRowIndex 起始行号（从0开始）
     * @param endRowIndex   结束行号（从0开始,包括此行）
     * @param startCellName 起始列名（从A开始）
     * @param endCellName   结束列名（从A开始,包括此列）
     * @return 表格数据
     */
    public static ExcelData readExcelData(File file, int startRowIndex, int endRowIndex, String startCellName, String endCellName) {
        return readExcelData(file, null, startRowIndex, endRowIndex, startCellName, endCellName);
    }

    /**
     * 读取表格指定行列范围内的数据
     *
     * @param file          文件
     * @param sheetName     sheet名称
     * @param startRowIndex 起始行号（从0开始）
     * @param endRowIndex   结束行号（从0开始,包括此行）
     * @param startCellName 起始列名（从A开始）
     * @param endCellName   结束列名（从A开始,包括此列）
     * @return 表格数据
     */
    public static ExcelData readExcelData(File file, String sheetName, int startRowIndex, int endRowIndex, String startCellName, String endCellName) {
        int startCellIndex = ExcelUtil.colNameToIndex(startCellName + "0");
        int endCellIndex = ExcelUtil.colNameToIndex(endCellName + "0");
        return readExcelData(file, sheetName, startRowIndex, endRowIndex, startCellIndex, endCellIndex);
    }

    /**
     * 读取表格中指定单元格的数据
     *
     * @param file        文件
     * @param locationRef 单元格位置（例如：A1）
     * @return 表格数据
     */
    public static Object readCellValue(File file, String locationRef) {
        return readCellValue(file, null, locationRef);
    }

    /**
     * 读取表格中指定单元格的数据
     *
     * @param file        文件
     * @param sheetName   sheet名称
     * @param locationRef 单元格位置（例如：A1）
     * @return 表格数据
     */
    public static Object readCellValue(File file, String sheetName, String locationRef) {
        ExcelReader excelReader = null;
        if (StrUtil.isBlank(sheetName)) {
            //sheet名称为空，默认读取第一个sheet
            excelReader = ExcelUtil.getReader(file, 0);
        } else {
            //读取指定的sheet
            excelReader = ExcelUtil.getReader(file, sheetName);
        }
        CellLocation cellLocation = ExcelUtil.toLocation(locationRef);
        return excelReader.readCellValue(cellLocation.getX(), cellLocation.getY());
    }

    /**
     * 获取所有的sheet名称
     *
     * @param file excel文件
     * @return sheet名称列表
     */
    public static List<String> getSheetNames(File file) {
        ExcelReader excelReader = ExcelUtil.getReader(file, 0);
        return excelReader.getSheetNames();
    }


    /**
     * 读取指定行列范围的表格数据
     *
     * @param file           文件
     * @param sheetName      sheet名称
     * @param startRowIndex  起始行号（从0开始）
     * @param endRowIndex    结束行号（从0开始,包括此行）
     * @param startCellIndex 起始列号（从0开始）
     * @param endCellIndex   结束列号（从0开始,包括此列）
     * @return 表格数据
     */
    private static ExcelData readExcelData(File file, String sheetName, int startRowIndex, int endRowIndex, int startCellIndex, int endCellIndex) {
        ExcelData excelData = new ExcelData();
        ExcelReader excelReader = null;
        if (StrUtil.isBlank(sheetName)) {
            //sheet名称为空，默认读取第一个sheet
            excelReader = ExcelUtil.getReader(file, 0);
        } else {
            //读取指定的sheet
            excelReader = ExcelUtil.getReader(file, sheetName);
        }
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
