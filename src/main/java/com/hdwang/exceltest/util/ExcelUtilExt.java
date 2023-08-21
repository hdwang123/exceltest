package com.hdwang.exceltest.util;

import cn.hutool.poi.excel.ExcelWriter;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

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
        Cell cell = excelWriter.getCell(x, y);
        CellStyle cellStyleSrc = cell.getCellStyle();
        //必须新创建单元格样式，直接修改原单元格样式可能影响到其它单元格，因为样式可以复用的
        CellStyle cellStyleDest = excelWriter.createCellStyle(x, y);
        //原单元格样式不为空，先拷贝原单元格样式至新创建的单元格样式
        if (cellStyleSrc != null) {
            cellStyleDest.cloneStyleFrom(cellStyleSrc);
        }
        cellStyleDest.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyleDest.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
    }

    /**
     * 给Cell添加批注
     *
     * @param cell   单元格
     * @param value  批注内容
     * @param isXlsx 是否是xlsx格式的文档
     */
    public static void addCellComment(Cell cell, String value, boolean isXlsx) {
        Sheet sheet = cell.getSheet();
        cell.removeCellComment();
        Drawing drawing = sheet.createDrawingPatriarch();
        Comment comment;
        if (isXlsx) {
            // 创建批注
            comment = drawing.createCellComment(new XSSFClientAnchor(1, 1, 1, 1, 1, 1, 1, 1));
            // 输入批注信息
            comment.setString(new XSSFRichTextString(value));
            // 将批注添加到单元格对象中
        } else {
            // 创建批注
            comment = drawing.createCellComment(new HSSFClientAnchor(1, 1, 1, 1, (short) 1, 1, (short) 1, 1));
            // 输入批注信息
            comment.setString(new HSSFRichTextString(value));
            // 将批注添加到单元格对象中
        }
        cell.setCellComment(comment);
    }
}
