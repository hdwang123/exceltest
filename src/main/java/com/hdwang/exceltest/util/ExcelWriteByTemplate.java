package com.hdwang.exceltest.util;

import cn.hutool.core.util.ReUtil;
import cn.hutool.poi.excel.cell.CellLocation;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

/**
 * POI 根据模板写入工具类
 *
 * @author wanghuidong
 * 时间： 2025/2/6 9:16
 */
public class ExcelWriteByTemplate {

    public static void main(String[] args) {
        String filePath = "src/main/resources/tmp1.xlsx";
        String targetFilePath = "src/main/resources/target_tmp1.xlsx";
        Workbook workbook = read(filePath);
        Sheet sheet = workbook.getSheetAt(0);

        //参数替换
        Map<String, String> params = new HashMap<>();
        params.put("{age}", "30");
        params.put("{name}", "张三");
        replaceByParam(sheet, params);

        //位置替换
        Map<String, String> locationValues = new HashMap<>();
        locationValues.put("A2", "值1");
        locationValues.put("B2", "值2");
        replaceByLocation(sheet, locationValues);

        //动态插入行测试
        String startLocation = "B7";
        List<List<String>> dataList = new ArrayList<>();
        dataList.add(Arrays.asList("张三", "20", "男", "杭州人", "爱好"));
        dataList.add(Arrays.asList("李四", "20", "男", "杭州人", "爱好"));
        dataList.add(Arrays.asList("王五", "20", "男", "杭州人", "爱好"));
        dataList.add(Arrays.asList("小刘", "20", "男", "杭州人", "爱好"));
        dataList.add(Arrays.asList("小七", "20", "男", "杭州人", "爱好"));
        insertRows(sheet, startLocation, dataList);

        //参数替换
        params = new HashMap<>();
        params.put("{ageTotal}", "100");
        replaceByParam(sheet, params);

        //输出新文件
        writeFile(targetFilePath, workbook);
    }

    private static void insertRows(Sheet sheet, String startLocation, List<List<String>> dataList) {
        CellLocation cellLocation = toLocation(startLocation);
        int startRowIndex = cellLocation.getY();
        int startCellIndex = cellLocation.getX();
        int endRow = sheet.getLastRowNum();
        int rowCount = dataList.size();

        Row startRow = sheet.getRow(startRowIndex);
        CellStyle rowStyle = startRow.getRowStyle();
        List<Cell> cells = new ArrayList<>();
        for (Cell cell : startRow) {
            cells.add(cell);
        }

        // 起始行小于等于最后行，是插入，需要移动后续行
        if (startRowIndex <= endRow) {
            // 移动后续行
            sheet.shiftRows(startRowIndex + 1, endRow, rowCount - 1);
        }
        // 创建行
        int dataIndex = 0;
        for (int rowIndex = startRowIndex; rowIndex < startRowIndex + rowCount; rowIndex++, dataIndex++) {
            // 读取占位行
            Row row = null;
            if (rowIndex == startRowIndex) {
                row = sheet.getRow(rowIndex);
            } else {
                //创建行
                row = sheet.createRow(rowIndex);
                row.setRowStyle(rowStyle);
                //创建列
                for (Cell cell : cells) {
                    row.createCell(cell.getColumnIndex()).setCellStyle(cell.getCellStyle());
                }
            }

            //赋值
            List<String> rowData = dataList.get(dataIndex);
            int cellIndex = startCellIndex;
            for (String cellVal : rowData) {
                row.getCell(cellIndex).setCellValue(cellVal);
                cellIndex++;
            }
        }
    }

    private static void writeFile(String targetFilePath, Workbook workbook) {
        try {
            OutputStream outputStream = new FileOutputStream(targetFilePath);
            workbook.write(outputStream);
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    /**
     * 读取Excel文件
     *
     * @param filePath 文件路径
     * @return 工作簿
     */
    public static Workbook read(String filePath) {
        try {
            // 读取Excel模板文件
            InputStream file = new FileInputStream(filePath);
            Workbook workbook = null;
            if (filePath.endsWith(".xlsx")) {
                workbook = new XSSFWorkbook(file);
            } else {
                workbook = new HSSFWorkbook(file);
            }
            return workbook;
        } catch (Exception ex) {
            ex.printStackTrace();
        }
        return null;
    }


    /**
     * 使用参数替换
     *
     * @param sheet  工作表
     * @param params 参数，格式：{"{age}":"30","{name}":"张三"}
     */
    public static void replaceByParam(Sheet sheet, Map<String, String> params) {
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() == CellType.STRING) {
                    String cellValue = cell.getStringCellValue();
                    // 检查是否包含占位符，如 {age}
                    if (params.containsKey(cellValue)) {
                        cell.setCellValue(params.get(cellValue));
                    }
                }
            }
        }
    }

    /**
     * 使用位置替换
     *
     * @param sheet          工作表
     * @param locationValues 位置值，格式：{ "A1":"20","B2":"张三"}
     */
    public static void replaceByLocation(Sheet sheet, Map<String, String> locationValues) {
        locationValues.forEach((location, value) -> {
            CellLocation cellLocation = toLocation(location);
            Row row = sheet.getRow(cellLocation.getY());
            if (row != null) {
                Cell cell = row.getCell(cellLocation.getX());
                if (cell != null) {
                    cell.setCellValue(value);
                }
            }
        });
    }


    /**
     * 将Excel中地址标识符（例如A11，B5）等转换为行列表示<br>
     * 例如：A11 -》 x:0,y:10，B5-》x:1,y:4
     *
     * @param locationRef 单元格地址标识符，例如A11，B5
     * @return 坐标点，x表示列号，从0开始，y表示行号，从0开始
     * @since 5.1.4
     */
    public static CellLocation toLocation(String locationRef) {
        final int x = colNameToIndex(locationRef);
        final int y = ReUtil.getFirstNumber(locationRef) - 1;
        return new CellLocation(x, y);
    }

    /**
     * 根据表元的列名转换为列号
     *
     * @param colName 列名, 从A开始
     * @return A1-》0; B1-》1...AA1-》26
     * @since 4.1.20
     */
    public static int colNameToIndex(String colName) {
        int length = colName.length();
        char c;
        int index = -1;
        for (int i = 0; i < length; i++) {
            c = Character.toUpperCase(colName.charAt(i));
            if (Character.isDigit(c)) {
                break;// 确定指定的char值是否为数字
            }
            index = (index + 1) * 26 + (int) c - 'A';
        }
        return index;
    }
}
