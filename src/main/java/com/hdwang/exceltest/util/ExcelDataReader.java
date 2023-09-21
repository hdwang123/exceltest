package com.hdwang.exceltest.util;

import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.cell.CellHandler;
import cn.hutool.poi.excel.cell.CellLocation;
import com.hdwang.exceltest.model.CellData;
import com.hdwang.exceltest.model.ExcelData;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;

import java.io.File;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * 表格数据工具类
 *
 * @author wanghuidong
 * @date 2022/1/27 16:12
 */
public class ExcelDataReader {

    /**
     * 读取表格指定行列范围内的数据,默认读取第一个sheet里的内容
     *
     * @param file          文件
     * @param startRowIndex 起始行号（从0开始）
     * @param endRowIndex   结束行号（从0开始,包括此行）
     * @param startColIndex 起始列号（从0开始）
     * @param endColIndex   结束列号（从0开始,包括此列）
     * @return 表格数据
     */
    public static ExcelData readExcelData(File file, int startRowIndex, int endRowIndex, int startColIndex, int endColIndex) {
        return readExcelData(file, null, startRowIndex, endRowIndex, startColIndex, endColIndex);
    }

    /**
     * 读取表格指定行列范围内的数据,默认读取第一个sheet里的内容
     *
     * @param file          文件
     * @param startRowIndex 起始行号（从0开始）
     * @param endRowIndex   结束行号（从0开始,包括此行）
     * @param startColName  起始列名（从A开始）
     * @param endColName    结束列名（从A开始,包括此列）
     * @return 表格数据
     */
    public static ExcelData readExcelData(File file, int startRowIndex, int endRowIndex, String startColName, String endColName) {
        return readExcelData(file, null, startRowIndex, endRowIndex, startColName, endColName);
    }

    /**
     * 读取表格指定行列范围内的数据
     *
     * @param file          文件
     * @param sheetName     sheet名称
     * @param startRowIndex 起始行号（从0开始）
     * @param endRowIndex   结束行号（从0开始,包括此行）
     * @param startColName  起始列名（从A开始）
     * @param endColName    结束列名（从A开始,包括此列）
     * @return 表格数据
     */
    public static ExcelData readExcelData(File file, String sheetName, int startRowIndex, int endRowIndex, String startColName, String endColName) {
        int startColIndex = ExcelUtil.colNameToIndex(startColName + "0");
        int endColIndex = ExcelUtil.colNameToIndex(endColName + "0");
        return readExcelData(file, sheetName, startRowIndex, endRowIndex, startColIndex, endColIndex);
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
        try {
            if (StrUtil.isBlank(sheetName)) {
                //sheet名称为空，默认读取第一个sheet
                excelReader = ExcelUtil.getReader(file, 0);
            } else {
                //读取指定的sheet
                excelReader = ExcelUtil.getReader(file, sheetName);
            }
            CellLocation cellLocation = ExcelUtil.toLocation(locationRef);
            return excelReader.readCellValue(cellLocation.getX(), cellLocation.getY());
        } finally {
            if (excelReader != null) {
                excelReader.close();
            }
        }
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
     * @param file          文件
     * @param sheetName     sheet名称
     * @param startRowIndex 起始行号（从0开始）
     * @param endRowIndex   结束行号（从0开始,包括此行）
     * @param startColIndex 起始列号（从0开始）
     * @param endColIndex   结束列号（从0开始,包括此列）
     * @return 表格数据
     */
    public static ExcelData readExcelData(File file, String sheetName, int startRowIndex, int endRowIndex, int startColIndex, int endColIndex) {
        ExcelData excelData = new ExcelData();
        ExcelReader excelReader = null;
        try {
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
                    if (cell.getColumnIndex() < startColIndex || cell.getColumnIndex() > endColIndex) {
                        //列号不在范围内跳过
                        return;
                    }

                    //值格式转换
                    if (value != null) {
                        String valueStr = value.toString().trim();
                        if (StrUtil.isNotBlank(valueStr)) {
                            //格式转换
                            if (isNumericPercent(cell)) {
                                //数字格式的百分比：转成字符串存入数据库,否则存入小数无法直接判断出是否是百分号
                                value = NumberUtil.decimalToPercent(valueStr);
                            } else if (isStrPercent(cell, valueStr)) {
                                //字符串格式的百分比：去掉逗号、空格等字符
                                value = valueStr.replaceAll("[,%]", "") + "%";
                            } else if (NumberUtil.isNumber(valueStr.replaceAll("[,]", ""))) {
                                //数值类型转成原始数据格式（非科学计数法等）
                                value = NumberUtil.toBigDecimal(valueStr.replaceAll("[,]", "")).toPlainString();
                            } else if (isFormattedDate(cell)) {
                                //日期格式转换成字符串
                                String format = ExcelDateUtil.getJavaDateFormat(cell.getCellStyle().getDataFormatString());
                                value = cn.hutool.core.date.DateUtil.format((Date) value, format);
                            }
                        }
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
        } finally {
            if (excelReader != null) {
                excelReader.close();
            }
        }
        return excelData;
    }

    /**
     * 是否是格式化日期字符串
     *
     * @param cell 单元格
     * @return
     */
    private static boolean isFormattedDate(Cell cell) {
        if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
            CellStyle cellStyle = cell.getCellStyle();
            if (cellStyle != null) {
                return StrUtil.isNotBlank(cellStyle.getDataFormatString());
            }
        }
        return false;
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

    /**
     * 判断是否是数字百分比
     *
     * @param cell 单元格
     * @return 是否是数字百分比
     */
    private static boolean isNumericPercent(Cell cell) {
        if (cell.getCellType() == CellType.NUMERIC || cell.getCellType() == CellType.FORMULA) {
            String dataFormatStr = cell.getCellStyle().getDataFormatString();
            if (dataFormatStr != null && dataFormatStr.contains("%")) {
                return true;
            }
        }
        return false;
    }

    /**
     * 判断是否是字符串百分比
     *
     * @param cell     单元格
     * @param valueStr 值
     * @return 是否是字符串百分比
     */
    private static boolean isStrPercent(Cell cell, String valueStr) {
        if (cell.getCellType() == CellType.STRING) {
            if (valueStr.contains("%") && NumberUtil.isNumber(valueStr.replaceAll("[,%]", ""))) {
                return true;
            }
        }
        return false;
    }

}
