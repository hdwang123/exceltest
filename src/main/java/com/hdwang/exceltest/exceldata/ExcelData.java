package com.hdwang.exceltest.exceldata;

import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.cell.CellLocation;
import org.apache.commons.collections4.CollectionUtils;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * 表格数据对象
 */
public class ExcelData {

    /**
     * 两层list表示的行列数据
     */
    private List<List<CellData>> rowDataList;

    /**
     * Map表示的单元格数据
     */
    private Map<String, CellData> cellDataMap;


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

    /**
     * 转换为bean对象列表
     *
     * @param tClass bean类型
     * @param <T>
     * @return bean对象列表
     */
    public <T> List<T> toBeanList(Class<T> tClass) {
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
