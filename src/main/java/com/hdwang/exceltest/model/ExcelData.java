package com.hdwang.exceltest.model;

import cn.hutool.core.util.NumberUtil;
import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.cell.CellLocation;
import org.apache.commons.collections4.CollectionUtils;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

/**
 * 表格数据对象
 *
 * @author wanghuidong
 * @date 2022/1/27 16:12
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

    /**
     * 表格列偏移量
     */
    private int tableColOffset;

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

    public int getTableColOffset() {
        return tableColOffset;
    }

    public void setTableColOffset(int tableColOffset) {
        this.tableColOffset = tableColOffset;
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
     * @param locationRef 单元格位置（例：A1）
     * @return 单元格数据
     */
    public CellData getCellData(String locationRef) {
        CellLocation cellLocation = ExcelUtil.toLocation(locationRef);
        return getCellData(cellLocation.getY(), cellLocation.getX());
    }

    /**
     * 获取指定单元格值
     *
     * @param rowIndex 行号（从0开始）
     * @param colIndex 列号（从0开始）
     * @return 单元格值
     */
    public Object getCellValue(int rowIndex, int colIndex) {
        String key = rowIndex + "_" + colIndex;
        CellData cellData = cellDataMap.get(key);
        if (cellData != null) {
            return cellData.getValue();
        }
        return null;
    }

    /**
     * 获取指定单元格值
     *
     * @param locationRef 单元格位置（例：A1）
     * @return 单元格值
     */
    public Object getCellValue(String locationRef) {
        CellLocation cellLocation = ExcelUtil.toLocation(locationRef);
        return getCellValue(cellLocation.getY(), cellLocation.getX());
    }


    /**
     * 获取指定单元格值并去掉百分号
     *
     * @param locationRef 单元格位置（例：A1）
     * @return 单元格值
     */
    public Object getCellValueWithoutPercent(String locationRef) {
        CellLocation cellLocation = ExcelUtil.toLocation(locationRef);
        Object cellValue = getCellValue(cellLocation.getY(), cellLocation.getX());
        if (cellValue != null) {
            String value = String.valueOf(cellValue).replaceAll("%", "");
            return StrUtil.isNotBlank(value) ? value : null;
        }
        return null;
    }


    /**
     * 获取合并单元格的值
     *
     * @param excelData        单元格数据对象
     * @param startLocationRef 单元格位置（例：A1）
     * @param endLocationRef   单元格位置（例：A2）
     * @return 单元格数据
     */
    public double getMergedRegionValue(ExcelData excelData, String startLocationRef, String endLocationRef) {
        CellLocation startLocation = ExcelUtil.toLocation(startLocationRef);
        CellLocation endLocation = ExcelUtil.toLocation(endLocationRef);
        int startRowIndex = startLocation.getY();
        int endRowIndex = endLocation.getY();
        int startCellIndex = startLocation.getX();
        int endCellIndex = endLocation.getX();
        if (startLocation.getY() > endLocation.getY()) {
            startRowIndex = endLocation.getY();
            endRowIndex = startLocation.getY();
        }
        if (startLocation.getX() > endLocation.getX()) {
            startCellIndex = endLocation.getX();
            endCellIndex = startLocation.getX();
        }
        BigDecimal sum = new BigDecimal(0);
        for (int r = startRowIndex; r <= endRowIndex; r++) {
            for (int c = startCellIndex; c <= endCellIndex; c++) {
                CellData data = excelData.getCellData(r, c);
                String valueStr;
                if (data != null) {
                    //转成字符串精确计算小数等
                    valueStr = data.getValue() == null ? "0" : String.valueOf(data.getValue());
                    if (StrUtil.isBlank(valueStr)) {
                        valueStr = "0";
                    }
                    if (NumberUtil.isNumber(valueStr)) {
                        BigDecimal value = NumberUtil.toBigDecimal(valueStr);
                        sum = sum.add(value);
                    }
                }
            }
        }
        return sum.doubleValue();
    }


    /**
     * 转换为bean对象列表
     * 注：默认列值将截去空格
     *
     * @param tClass bean类型
     * @param <T>
     * @return bean对象列表
     */
    public <T> List<T> toBeanList(Class<T> tClass) {
        return toBeanList(tClass, true, 100);
    }

    /**
     * 转换为bean对象列表
     *
     * @param tClass           bean类型
     * @param <T>
     * @param trim             列值是否截去空格
     * @param defaultMaxValLen 默认列值最大长度
     * @return bean对象列表
     */
    public <T> List<T> toBeanList(Class<T> tClass, boolean trim, int defaultMaxValLen) {
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
                        final int i = index + tableColOffset;
                        CellData cellData = rowData.stream().filter(x -> x.getCellIndex() == i).findFirst().orElse(null);
                        if (cellData != null && cellData.getValue() != null) {
                            Object value = cellData.getValue();
                            if (field.getType().getName().equals(String.class.getName())) {
                                String valueStr = String.valueOf(value);

                                //移除备注等无意义字符串
                                if (isNoNeedInfo(valueStr)) {
                                    valueStr = null;
                                } else {

                                    //截去首尾空白
                                    if (trim) {
                                        valueStr = StrUtil.trim(valueStr);
                                    }

                                    //全局替换换行符为:### , 解决大数据抽取不能带换行符问题
                                    valueStr = valueStr.replaceAll("[\r\n]", "#@#");

                                    //列值高级截去
                                    if (field.isAnnotationPresent(ColValTrimmer.class)) {
                                        ColValTrimmer colValTrimmer = field.getAnnotation(ColValTrimmer.class);
                                        int maxLen = colValTrimmer.maxLength();
                                        if (maxLen != -1 && valueStr.length() > maxLen) {
                                            valueStr = valueStr.substring(0, maxLen);
                                        }
                                        String[] regexes = colValTrimmer.trimRegex();
                                        for (String regex : regexes) {
                                            valueStr = valueStr.replaceAll(regex, "");
                                        }
                                    } else if (valueStr.length() > defaultMaxValLen) {
                                        //默认列值截短
                                        valueStr = valueStr.substring(0, defaultMaxValLen);
                                    }
                                }

                                //重新赋值回去
                                value = valueStr;
                            }
                            field.setAccessible(true);
                            field.set(bean, value);
                        }
                    }
                    //设置行号到bean里去
                    for (Field parentField : tClass.getSuperclass().getDeclaredFields()) {
                        if (parentField.getName().equals("rowNo")) {
                            parentField.setAccessible(true);
                            parentField.set(bean, rowData.get(0).getRowIndex() + 1);
                            break;
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

    /**
     * 判断是否是不需要的信息
     *
     * @param valueStr 值
     * @return 是否是不需要的信息
     */
    private boolean isNoNeedInfo(String valueStr) {
        String valueTrim = valueStr.trim();
        // 移除无意义字符
        List<String> meaninglessValues = Arrays.asList("/", "-", "不适用");
        if (meaninglessValues.contains(valueTrim)) {
            return true;
        }
        // 移除注释等内容
        if (valueStr.startsWith("注") && valueStr.length() > 100 || valueStr.startsWith("填报说明")) {
            return true;
        }
        return false;
    }

}
