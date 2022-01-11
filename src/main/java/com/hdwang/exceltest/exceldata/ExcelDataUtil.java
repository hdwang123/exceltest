package com.hdwang.exceltest.exceldata;

import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.cell.CellHandler;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.Cell;

import java.io.File;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * 表格数据工具类
 */
public class ExcelDataUtil {

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
    public static <T> ExcelData<T> readExcelData(File templateFile, int startRowIndex, int endRowIndex, int startCellIndex, int endCellIndex, Class<T> tClass) {
        ExcelData<T> excelData = new ExcelData<>();
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
        //转为为bean结构
        excelData.setBeanList(convertExcelDataToBeanList(rowDataList, tClass));
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


    /**
     * 转换表格数据为bean对象列表
     *
     * @param rowDataList 表格数据
     * @param tClass      bean类型s
     * @param <T>
     * @return bean对象列表
     */
    private static <T> List<T> convertExcelDataToBeanList(List<List<CellData>> rowDataList, Class<T> tClass) {
        if (CollectionUtils.isEmpty(rowDataList)) {
            return new ArrayList<>();
        }
        List<T> beanList = new ArrayList<>();
        for (List<CellData> rowData : rowDataList) {
            try {
                //实例化bean对象
                T bean = tClass.newInstance();
                //遍历字段并赋值
                Field[] fields = tClass.getDeclaredFields();
                for (Field field : fields) {
                    if (field.isAnnotationPresent(ColIndex.class)) {
                        ColIndex colIndex = field.getAnnotation(ColIndex.class);
                        int index = colIndex.index();
                        String name = colIndex.name();
                        if (index != -1) {
                            //do nothing
                        } else if (!"".equals(name)) {
                            //列名转索引号（补0为了适应下述方法）
                            index = ExcelUtil.colNameToIndex(name + "0");
                        } else {
                            throw new RuntimeException("对象属性上的ColIndex注解必须设置值");
                        }
                        //从行数据中找到指定单元格数据给字段赋值
                        final int i = index;
                        CellData cellData = rowData.stream().filter(x -> x.getCellIndex() == i).findFirst().orElse(null);
                        if (cellData != null) {
                            Object value = cellData.getValue();
                            if (field.getType().getName().equals(String.class.getName())) {
                                value = String.valueOf(value);
                            }
                            field.setAccessible(true);
                            field.set(bean, value);
                        }
                    }
                }
                beanList.add(bean);
            } catch (Exception ex) {
                throw new RuntimeException("实例化对象失败：" + ex.getMessage(), ex);
            }
        }
        return beanList;
    }


}
