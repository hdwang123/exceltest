package com.hdwang.exceltest.util;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

/**
 * excel格式转换工具类
 *
 * @author wanghuidong
 * 时间： 2023/9/11 11:12
 */
@Slf4j
public class ExcelFmtConvert {

    /**
     * Excel格式从xls转换成xlsx格式
     *
     * @param xlsInputStream   xls格式的输入流
     * @param xlsxOutputStream xlsx格式的输出流
     */
    public static void convertXlsToXlsxByStream(InputStream xlsInputStream, OutputStream xlsxOutputStream) {
        try {
            HSSFWorkbook oldWorkbook = new HSSFWorkbook(xlsInputStream);
            XSSFWorkbook newWorkbook = new XSSFWorkbook();

            for (int i = 0; i < oldWorkbook.getNumberOfSheets(); i++) {
                HSSFSheet oldSheet = oldWorkbook.getSheetAt(i);
                XSSFSheet newSheet = newWorkbook.createSheet(oldSheet.getSheetName());

                // 遍历行，创建行
                for (int j = 0; j <= oldSheet.getLastRowNum(); j++) {
                    HSSFRow oldRow = oldSheet.getRow(j);
                    XSSFRow newRow = newSheet.createRow(j);
                    // 复制行高
                    newRow.setHeight(oldRow.getHeight());

                    if (oldRow != null) {
                        //遍历列，复制单元格
                        for (int k = 0; k < oldRow.getLastCellNum(); k++) {
                            HSSFCell oldCell = oldRow.getCell(k);
                            XSSFCell newCell = newRow.createCell(k);
                            if (oldCell != null) {
                                try {
                                    setCellValue(newCell, oldCell);
                                    copyCellStyle(newWorkbook, newCell, oldWorkbook, oldCell);
                                } catch (Exception ex) {
                                    log.warn("单元格拷贝异常:", ex);
                                }
                            }
                        }
                    }
                }

                // 复制单元格合并信息
                List<CellRangeAddress> mergedRegions = oldSheet.getMergedRegions();
                for (CellRangeAddress mergedRegion : mergedRegions) {
                    CellRangeAddress targetMergedRegion = new CellRangeAddress(
                            mergedRegion.getFirstRow(),
                            mergedRegion.getLastRow(),
                            mergedRegion.getFirstColumn(),
                            mergedRegion.getLastColumn()
                    );
                    newSheet.addMergedRegion(targetMergedRegion);
                }

                // 复制列宽
                int columnCount = 0;
                if (newSheet.getRow(0) != null) {
                    // 假设第一行包含所有列，根据第一行的列数获取列数
                    columnCount = newSheet.getRow(0).getLastCellNum();
                }
                for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                    newSheet.setColumnWidth(columnIndex, oldSheet.getColumnWidth(columnIndex));
                }
            }

            newWorkbook.write(xlsxOutputStream);
            oldWorkbook.close();
            newWorkbook.close();
        } catch (Exception e) {
            log.error("excel格式转换(xls->xlsx)异常:", e);
        }
    }

    private static void setCellValue(XSSFCell newCell, HSSFCell oldCell) {
        if (oldCell == null) {
            return;
        }
        switch (oldCell.getCellType()) {
            case STRING:
                newCell.setCellValue(oldCell.getStringCellValue());
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(oldCell)) {
                    newCell.setCellValue(oldCell.getDateCellValue());
                } else {
                    newCell.setCellValue(oldCell.getNumericCellValue());
                }
                break;
            case BOOLEAN:
                newCell.setCellValue(oldCell.getBooleanCellValue());
                break;
            case FORMULA:
                newCell.setCellValue(oldCell.getCellFormula());
                break;
            default:
        }
    }

    private static void copyCellStyle(XSSFWorkbook xssfWorkbook, XSSFCell newCell, HSSFWorkbook hssfWorkbook, HSSFCell oldCell) {
        HSSFCellStyle oldCellStyle = oldCell.getCellStyle();

        // 创建一个XSSFCellStyle（新Excel格式）
        XSSFCellStyle newCellStyle = xssfWorkbook.createCellStyle();

        // 复制对齐方式
        newCellStyle.setAlignment(oldCellStyle.getAlignment());
        newCellStyle.setVerticalAlignment(oldCellStyle.getVerticalAlignment());

        // 复制字体属性
        XSSFFont newFont = xssfWorkbook.createFont();
        HSSFFont oldFont = oldCellStyle.getFont(hssfWorkbook);
        newFont.setFontName(oldFont.getFontName());
        newFont.setFontHeightInPoints(oldFont.getFontHeightInPoints());
        newFont.setColor(oldFont.getColor());
        newCellStyle.setFont(newFont);

        // 复制填充颜色
        newCellStyle.setFillPattern(oldCellStyle.getFillPattern());
        newCellStyle.setFillForegroundColor(oldCellStyle.getFillForegroundColor());
        newCellStyle.setFillBackgroundColor(oldCellStyle.getFillBackgroundColor());

        // 复制数据格式
        newCellStyle.setDataFormat(oldCellStyle.getDataFormat());

        newCell.setCellStyle(newCellStyle);
    }
}
