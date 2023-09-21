package com.hdwang.exceltest.util;

import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hpsf.CustomProperties;
import org.apache.poi.hpsf.DocumentSummaryInformation;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Excel工作类扩展
 *
 * @author wanghuidong
 * 时间： 2022/10/10 18:58
 */
@Slf4j
public class ExcelUtilExt {

    /**
     * 给单元格标黄
     *
     * @param excelWriter excel写对象
     * @param x           列号
     * @param y           行号
     */
    public static void markCellYellow(ExcelWriter excelWriter, int x, int y) {
        setCellBgColor(excelWriter, x, y, IndexedColors.YELLOW.getIndex());
    }

    /**
     * 给单元格标粉红
     *
     * @param excelWriter excel写对象
     * @param x           列号
     * @param y           行号
     */
    public static void markCellPink(ExcelWriter excelWriter, int x, int y) {
        setCellBgColor(excelWriter, x, y, IndexedColors.PINK.getIndex());
    }

    /**
     * 设置单元格背景色
     *
     * @param excelWriter excel写对象
     * @param x           列号
     * @param y           行号
     * @param color       颜色
     */
    public static void setCellBgColor(ExcelWriter excelWriter, int x, int y, short color) {
        Cell cell = excelWriter.getCell(x, y);
        CellStyle cellStyleSrc = cell.getCellStyle();
        //必须新创建单元格样式，直接修改原单元格样式可能影响到其它单元格，因为样式可以复用的
        CellStyle cellStyleDest = excelWriter.createCellStyle(x, y);
        //原单元格样式不为空，先拷贝原单元格样式至新创建的单元格样式
        if (cellStyleSrc != null) {
            cellStyleDest.cloneStyleFrom(cellStyleSrc);
        }
        cellStyleDest.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyleDest.setFillForegroundColor(color);
    }

    /**
     * 给Cell添加批注
     *
     * @param cell   单元格
     * @param value  批注内容
     * @param isXlsx 是否是xlsx格式的文档
     * @param append 是否追加批注
     */
    public static void addCellComment(Cell cell, String value, boolean isXlsx, boolean append) {
        Sheet sheet = cell.getSheet();
        String oldCellComment = StrUtil.EMPTY;
        if (cell.getCellComment() != null && cell.getCellComment().getString() != null) {
            oldCellComment = cell.getCellComment().getString().getString();
            oldCellComment = oldCellComment == null ? StrUtil.EMPTY : oldCellComment;
        }
        if (append && StrUtil.isNotBlank(oldCellComment)) {
            value = oldCellComment + ";" + value;
        }
        cell.removeCellComment();
        Drawing drawing = sheet.createDrawingPatriarch();
        Comment comment;
        if (isXlsx) {
            comment = drawing.createCellComment(new XSSFClientAnchor(1, 1, 1, 1, 1, 1, 1, 1));

            comment.setString(new XSSFRichTextString(value));
        } else {
            comment = drawing.createCellComment(new HSSFClientAnchor(1, 1, 1, 1, (short) 1, 1, (short) 1, 1));
            comment.setString(new HSSFRichTextString(value));
        }
        cell.setCellComment(comment);
    }

    /**
     * 获取excel文档自定义属性
     *
     * @param excelFilePath excel文件路径
     * @param propertyKey   属性名
     * @return 属性值
     */
    public static String getCustomPropVal(String excelFilePath, String propertyKey) {
        String propertyValue = null;
        ExcelReader reader = null;
        try {
            reader = ExcelUtil.getReader(excelFilePath);
            Workbook workbook = reader.getWorkbook();
            if (workbook instanceof HSSFWorkbook) {
                DocumentSummaryInformation summaryInformation = ((HSSFWorkbook) workbook).getDocumentSummaryInformation();
                CustomProperties customProperties = summaryInformation.getCustomProperties();
                if (customProperties.containsKey(propertyKey)) {
                    propertyValue = customProperties.get(propertyKey).toString();
                }
            } else if (workbook instanceof XSSFWorkbook) {
                // 获取文档属性
                XSSFWorkbook xssfWorkbook = (XSSFWorkbook) workbook;
                org.apache.poi.ooxml.POIXMLProperties.CustomProperties customProperties = xssfWorkbook.getProperties().getCustomProperties();
                if (customProperties.contains(propertyKey)) {
                    propertyValue = customProperties.getProperty(propertyKey).getLpwstr();
                }
            } else {
                log.error("不支持的 Excel 文件格式");
            }

            // 关闭工作簿
            reader.close();
        } catch (Exception ex) {
            log.error("获取excel文档自定义属性失败：", ex);
        } finally {
            try {
                // 关闭工作簿
                if (reader != null) {
                    reader.close();
                }
            } catch (Exception ex) {
                log.error("关闭ExcelWriter异常：", ex);
            }
        }
        return propertyValue;
    }

    /**
     * 添加excel文档自定义属性
     *
     * @param excelFilePath excel文件路径
     * @param propertyKey   属性名
     * @return 属性值
     */
    public static void putCustomProp(String excelFilePath, String propertyKey, String propertyVal) {
        ExcelWriter writer = null;
        try {
            writer = ExcelUtil.getWriter(excelFilePath);
            Workbook workbook = writer.getWorkbook();
            ExcelUtil.getWriter(excelFilePath).close();
            if (workbook instanceof HSSFWorkbook) {
                DocumentSummaryInformation summaryInformation = ((HSSFWorkbook) workbook).getDocumentSummaryInformation();
                CustomProperties customProperties = summaryInformation.getCustomProperties();
                customProperties.put(propertyKey, propertyVal);
            } else if (workbook instanceof XSSFWorkbook) {
                // 获取文档属性
                XSSFWorkbook xssfWorkbook = (XSSFWorkbook) workbook;
                org.apache.poi.ooxml.POIXMLProperties.CustomProperties customProperties = xssfWorkbook.getProperties().getCustomProperties();
                customProperties.addProperty(propertyKey, propertyVal);
            } else {
                log.error("不支持的 Excel 文件格式");
            }
        } catch (Exception ex) {
            log.error("添加excel文档自定义属性失败：", ex);
        } finally {
            try {
                // 关闭工作簿
                if (writer != null) {
                    writer.close();
                }
            } catch (Exception ex) {
                log.error("关闭ExcelWriter异常：", ex);
            }
        }
    }
}
