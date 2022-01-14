package com.hdwang.exceltest;

import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;
import com.hdwang.exceltest.exceldata.ExcelData;
import com.hdwang.exceltest.exceldata.ExcelDataReader;
import com.hdwang.exceltest.validate.ExcelValidator;
import com.hdwang.exceltest.validate.validator.EqualValidator;
import com.hdwang.exceltest.validate.validator.NotNullValidator;
import com.hdwang.exceltest.validate.ValidateResult;
import com.hdwang.exceltest.validate.validator.RegexpValidator;
import com.hdwang.exceltest.validate.validator.SumValidator;
import org.apache.poi.ss.usermodel.*;


import java.io.File;
import java.util.*;

public class Main {
    public static void main(String[] args) {
        try {
            File templateFile = new File("C:\\Users\\hdwang\\Desktop\\test.xlsx");
            File outputFile = new File("C:\\Users\\hdwang\\Desktop\\test2.xlsx");
            //根据excel模板写值
            produceExcelByTemplate(templateFile, outputFile);

            long startTime = System.currentTimeMillis();
            //读取指定行列范围内的数据
            ExcelData reportExcelData = ExcelDataReader.readExcelData(templateFile, 2, Integer.MAX_VALUE, "B", "G");
            System.out.println("rowDataList:" + reportExcelData.getRowDataList()); //获取行列表示的数据
            System.out.println("cellDataMap:" + reportExcelData.getCellDataMap()); //获取map表示的数据
            //转换为beanList
            System.out.println("beanList:" + reportExcelData.toBeanList(ZhenquanReport.class));
            //从读取的数据中获取指定单元格数据
            System.out.println("C3:" + reportExcelData.getCellData(2, "C"));
            System.out.println("C3:" + reportExcelData.getCellData("C3"));
            //直接读取Excel文件中某个单元格的数值
            System.out.println("C3:" + ExcelDataReader.readCellValue(templateFile, "C3"));
            System.out.println("测试2 B2:" + ExcelDataReader.readCellValue(templateFile, "测试2", "B2"));
            System.out.println("cost time:" + (System.currentTimeMillis() - startTime) + "ms");

            //指定行列范围内的所有单元格的非空校验
            System.out.println("==================非空校验======================");
            List<ValidateResult> results = ExcelValidator.validate(reportExcelData, 2, 5, "B", "E", new NotNullValidator());
            System.out.println(results);
            //指定单元格的非空校验
            ValidateResult result = ExcelValidator.validate(reportExcelData, 4, "E", new NotNullValidator());
            System.out.println(result);
            result = ExcelValidator.validate(reportExcelData, "A5", new NotNullValidator());
            System.out.println(result);
            //指定单元格的求和校验
            System.out.println("==================求和校验======================");
            result = ExcelValidator.validate(reportExcelData, "C7", new SumValidator("C3", "C6"));
            System.out.println(result);
            result = ExcelValidator.validate(reportExcelData, "F7", new SumValidator("F3", "F6"));
            System.out.println(result);
            result = ExcelValidator.validate(reportExcelData, "G5", new SumValidator("G3", "G4"));
            System.out.println(result);
            //单元格等值校验
            System.out.println("==================等值校验======================");
            result = ExcelValidator.validate(reportExcelData, "C7", new EqualValidator("3000"));
            System.out.println(result);
            result = ExcelValidator.validate(reportExcelData, "B6", new EqualValidator("净利"));
            System.out.println(result);
            //单元格格式校验（正则式校验）
            System.out.println("==================格式校验======================");
            result = ExcelValidator.validate(reportExcelData, "F4", new RegexpValidator("\\d+\\.\\d{2}", "格式不正确，请保留两位小数"));
            System.out.println(result);
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }


    /**
     * 根据模板文件产生Excel文件
     *
     * @param templateFile 模板文件
     * @param outputFile   输出文件
     */
    private static void produceExcelByTemplate(File templateFile, File outputFile) {
        //===========================读取测试=====================================
        ExcelReader excelReader = ExcelUtil.getReader(templateFile, 0);

        //读取sheet名
        List<String> sheetNames = excelReader.getSheetNames();
        System.out.println(sheetNames);

        //读取数据，按照行列方式读取所有数据
        List<List<Object>> rowObjects = excelReader.read();
        System.out.println(rowObjects);

        //读取指定行数据
        rowObjects = excelReader.read(2, Integer.MAX_VALUE, false);
        System.out.println(rowObjects);


        //读取数据，指定标题行和起始数据行
        List<Map<String, Object>> rowMaps = excelReader.read(1, 2, Integer.MAX_VALUE);
        System.out.println(rowMaps);

        //读取数据，指定标题行和起始数据行，转换为对象
        excelReader.addHeaderAlias("名称", "name");
        excelReader.addHeaderAlias("数值", "value");
        excelReader.addHeaderAlias("E", "value2");
        List<ZhenquanReport> reports = excelReader.read(1, 2, Integer.MAX_VALUE, ZhenquanReport.class);
        System.out.println(reports);

        //读取指定单元格
        String value = String.valueOf(excelReader.readCellValue(2, 2));
        System.out.println(value);

        //关闭
        excelReader.close();

        //===========================写入测试=====================================
        ExcelWriter excelWriter = new ExcelWriter(templateFile);
        excelWriter.writeCellValue(3, 2, "error"); //写入值，x=列、y=行号

        //设置单元格样式
        CellStyle cellStyle = excelWriter.createCellStyle(3, 2);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND); //设置背景色
        cellStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        cellStyle.setBorderBottom(BorderStyle.DASHED); //设置边框线条与颜色
        cellStyle.setBottomBorderColor(IndexedColors.PINK.getIndex());
        Font font = excelWriter.createFont();
        font.setColor(IndexedColors.RED.getIndex());
        cellStyle.setFont(font); //设置字体

        //设置输出文件路径
        excelWriter.setDestFile(outputFile);
        excelWriter.close();
    }


}
