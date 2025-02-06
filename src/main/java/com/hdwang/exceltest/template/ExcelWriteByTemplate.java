package com.hdwang.exceltest.template;

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
        String filePath = "src/main/resources/tmp1.xls";
        String targetFilePath = "src/main/resources/target_tmp1.xls";
        Workbook workbook = read(filePath);
        Sheet sheet = workbook.getSheetAt(0);

        //参数替换
        Map<String, String> params = new HashMap<>();
        params.put("{age}", "30");
        params.put("{name}", "张三");
        params.put("{ageTotal}", "100");
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

        //输出新文件
        writeFile(targetFilePath, workbook);
    }

    /**
     * 读取Excel文件
     *
     * @param filePath 文件路径
     * @return 工作簿
     */
    public static Workbook read(String filePath) {
        // 读取Excel模板文件
        try (InputStream file = new FileInputStream(filePath);) {
            Workbook workbook = null;
            if (filePath.endsWith(".xlsx")) {
                workbook = new XSSFWorkbook(file);
            } else {
                workbook = new HSSFWorkbook(file);
            }
            return workbook;
        } catch (Exception ex) {
            throw new RuntimeException("读取文件异常", ex);
        }
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
     * 输出Excel文件
     *
     * @param targetFilePath 目标文件路径
     * @param workbook       工作簿
     */
    public static void writeFile(String targetFilePath, Workbook workbook) {
        try (OutputStream outputStream = new FileOutputStream(targetFilePath);) {
            workbook.write(outputStream);
        } catch (Exception ex) {
            throw new RuntimeException("输出文件异常", ex);
        }
    }

    /**
     * 插入行,从指定位置开始插入数据，保留原占位行的格式
     *
     * @param sheet         工作表
     * @param startLocation 起始单元格
     * @param dataList      数据列表
     */
    public static void insertRows(Sheet sheet, String startLocation, List<List<String>> dataList) {
        CellLocation cellLocation = toLocation(startLocation);
        int startRowIndex = cellLocation.getY();
        int startCellIndex = cellLocation.getX();
        int endRow = sheet.getLastRowNum();
        int rowCount = dataList.size();

        // 读取占位行的格式
        Row startRow = sheet.getRow(startRowIndex);
        if (startRow == null) {
            startRow = sheet.createRow(startRowIndex);
        }
        CellStyle rowStyle = startRow.getRowStyle();
        List<Cell> cells = new ArrayList<>();
        for (Cell cell : startRow) {
            cells.add(cell);
        }

        // 起始行小于最后行，是插入，需要移动后续行
        if (startRowIndex < endRow) {
            // 移动后续行
            sheet.shiftRows(startRowIndex + 1, endRow, rowCount - 1);
        }
        // 创建行
        int dataIndex = 0;
        for (int rowIndex = startRowIndex; rowIndex < startRowIndex + rowCount; rowIndex++, dataIndex++) {
            // 读取占位行或创建行
            Row row = null;
            if (rowIndex == startRowIndex) {
                row = sheet.getRow(rowIndex);
                if (row == null) {
                    row = createRow(sheet, rowStyle, cells, rowIndex);
                }
            } else {
                row = createRow(sheet, rowStyle, cells, rowIndex);
            }

            //赋值
            List<String> rowData = dataList.get(dataIndex);
            int cellIndex = startCellIndex;
            for (String cellVal : rowData) {
                Cell cell = row.getCell(cellIndex);
                if (cell == null) {
                    cell = row.createCell(cellIndex);
                }
                cell.setCellValue(cellVal);
                cellIndex++;
            }
        }
    }

    /**
     * 创建行
     *
     * @param sheet    工作表
     * @param rowStyle 行样式
     * @param cells    单元格列表
     * @param rowIndex 行号
     * @return 行
     */
    private static Row createRow(Sheet sheet, CellStyle rowStyle, List<Cell> cells, int rowIndex) {
        //创建行
        Row row = sheet.createRow(rowIndex);
        row.setRowStyle(rowStyle);
        //创建列
        for (Cell cell : cells) {
            row.createCell(cell.getColumnIndex()).setCellStyle(cell.getCellStyle());
        }
        return row;
    }

    /**
     * 将Excel中地址标识符（例如A11，B5）等转换为行列表示<br>
     * 例如：A11 -》 x:0,y:10，B5-》x:1,y:4
     *
     * @param locationRef 单元格地址标识符，例如A11，B5
     * @return 坐标点，x表示列号，从0开始，y表示行号，从0开始
     * @since 5.1.4
     */
    private static CellLocation toLocation(String locationRef) {
        final int x = colNameToIndex(locationRef);
        final int y = Integer.parseInt(locationRef.replaceAll("[^0-9]", "")) - 1;
        return new CellLocation(x, y);
    }

    /**
     * 根据表元的列名转换为列号
     *
     * @param colName 列名, 从A开始
     * @return A1-》0; B1-》1...AA1-》26
     * @since 4.1.20
     */
    private static int colNameToIndex(String colName) {
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
