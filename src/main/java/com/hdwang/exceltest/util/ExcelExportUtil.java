//package com.hdwang.exceltest.util;
//
//import cn.hutool.core.bean.BeanUtil;
//import cn.hutool.core.collection.CollUtil;
//import cn.hutool.core.io.FileUtil;
//import cn.hutool.core.io.IoUtil;
//import cn.hutool.core.io.resource.ResourceUtil;
//import cn.hutool.core.net.URLEncodeUtil;
//import cn.hutool.core.text.CharPool;
//import cn.hutool.core.util.IdUtil;
//import cn.hutool.core.util.NumberUtil;
//import cn.hutool.core.util.ObjectUtil;
//import cn.hutool.core.util.StrUtil;
//import cn.hutool.poi.excel.BigExcelWriter;
//import cn.hutool.poi.excel.ExcelUtil;
//import cn.hutool.poi.excel.WorkbookUtil;
//import cn.hutool.poi.excel.cell.CellUtil;
//import cn.hutool.poi.excel.style.StyleUtil;
//import com.yc.cloud.common.model.util.ExcelBgColor;
//import com.yc.cloud.common.model.util.ExcelColumnConfig;
//import com.yc.cloud.common.model.util.ExcelMergeScope;
//import com.yc.cloud.common.model.util.ExcelWriteConfig;
//import com.yc.cloud.common.utils.ServletStaticUtil;
//import com.yc.cloud.common.utils.TmpFileUtil;
//import lombok.SneakyThrows;
//import lombok.extern.slf4j.Slf4j;
//import org.apache.poi.ss.usermodel.*;
//import org.apache.poi.xssf.streaming.SXSSFWorkbook;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//
//import javax.servlet.ServletOutputStream;
//import javax.servlet.http.HttpServletResponse;
//import java.io.BufferedOutputStream;
//import java.io.File;
//import java.util.Collection;
//import java.util.HashMap;
//import java.util.List;
//import java.util.Map;
//
//import static com.yc.cloud.common.constant.BaseConstant.ONE;
//import static com.yc.cloud.common.constant.BaseConstant.ZERO;
//import static com.yc.cloud.common.constant.UriConstant.TEMPORARY_URI;
//
///**
// * excel表格导出工具类
// *
// * @author 杨智杰
// * @since 2021/9/7 10:06
// */
//@Slf4j
//public class ExcelExportUtil {
//
//    public static final String XLSX = CharPool.DOT + "xlsx";
//
//    /**
//     * 模板导出
//     *
//     * @param templateName 模板文件名
//     * @param obj          导出内容
//     */
//    public static void downExcel(String templateName, Object obj) {
//        downExcel(FileUtil.file(ResourceUtil.getResource(templateName)), obj);
//    }
//
//    /**
//     * 模板导出
//     *
//     * @param template 模板文件
//     * @param obj      导出内容
//     */
//    public static void downExcel(File template, Object obj) {
//        downExcel(template, obj, IdUtil.simpleUUID() + XLSX);
//    }
//
//    /**
//     * 模板导出
//     *
//     * @param template 模板文件
//     * @param obj      导出内容
//     * @param fileName 导出文件名
//     */
//    public static void downExcel(File template, Object obj, String fileName) {
//        downExcel(template, obj, ServletStaticUtil.getResponse(), fileName);
//    }
//
//    /**
//     * 模板导出
//     *
//     * @param template 模板文件
//     * @param obj      导出内容
//     * @param response response
//     * @param fileName 导出文件名
//     */
//    @SneakyThrows
//    public static void downExcel(File template, Object obj, HttpServletResponse response, String fileName) {
//        if (!FileUtil.exist(template)) {
//            throw new Exception("未获取到模板文件!");
//        }
//        if (ObjectUtil.isNull(obj)) {
//            throw new Exception("导出数据不能为null!");
//        }
//        Map<String, Object> objectMap = BeanUtil.beanToMap(obj);
//        File copy = FileUtil.copy(template, FileUtil.file(TEMPORARY_URI + fileName), true);
//        Workbook book = WorkbookUtil.createBookForWriter(copy);
//        for (int i = ZERO; i < book.getNumberOfSheets(); i++) {
//            Sheet sheet = book.getSheetAt(i);
//            for (int j = ZERO; j < (sheet.getLastRowNum() + ONE); j++) {
//                Row row = sheet.getRow(j);
//                if (ObjectUtil.isNotNull(row)) {
//                    for (int k = ZERO; k < row.getLastCellNum(); k++) {
//                        int finalK = k;
//                        objectMap.keySet().stream().forEach(key -> {
//                            if (String.valueOf(row.getCell(finalK)).contains("${" + key + "}")) {
//                                Cell cell = row.getCell(finalK);
//                                String objValue = String.valueOf(objectMap.get(key));
//                                String cellValue = cell.getStringCellValue();
//                                cell.setCellValue(StrUtil.replace(cellValue, "${" + key + "}", objValue, true));
//                            }
//                        });
//                    }
//                }
//            }
//        }
//        BufferedOutputStream outputStream = FileUtil.getOutputStream(copy);
//        book.write(outputStream);
//        outputStream.flush();
//        book.close();
//        IoUtil.close(outputStream);
//        book = WorkbookUtil.createBookForWriter(copy);
//        response.setContentType("application/vnd.ms-excel;charset=utf-8");
//        response.setHeader("Content-Disposition", "attachment;filename=" + fileName);
//        ServletOutputStream out = response.getOutputStream();
//        book.write(out);
//        out.flush();
//        book.close();
//        IoUtil.close(out);
//    }
//
//
//    /**
//     * 导出excel
//     *
//     * @param data 数据集合
//     */
//    public static void downExcel(Collection data) {
//        downExcel(data, null);
//    }
//
//    /**
//     * 导出excel
//     *
//     * @param configs 配置
//     * @param data    数据集合
//     */
//    public static void downExcel(Collection data, List<ExcelColumnConfig> configs) {
//        downExcel(data, configs, IdUtil.simpleUUID() + XLSX, null, null);
//    }
//
//    /**
//     * 导出excel
//     *
//     * @param data     数据集合
//     * @param configs  配置
//     * @param fileName 文件名
//     */
//    public static void downExcel(Collection data, List<ExcelColumnConfig> configs, String fileName) {
//        downExcel(data, configs, fileName, null, null);
//    }
//
//    /**
//     * 导出excel
//     *
//     * @param data           数据集合
//     * @param configs        配置
//     * @param fileName       文件名
//     * @param mergeScopeList 合并范围
//     * @param bgColors       背景色
//     */
//    public static void downExcel(Collection data, List<ExcelColumnConfig> configs, String fileName, List<ExcelMergeScope> mergeScopeList, List<ExcelBgColor> bgColors) {
//        //定义目标文件
//        File destFile = TmpFileUtil.getTmpFile(fileName);
//
//        downExcel(data, configs, destFile, null, true, true, 0, mergeScopeList, bgColors);
//    }
//
//    /**
//     * 导出excel
//     *
//     * @param data             数据集合
//     * @param configs          配置
//     * @param excelWriteConfig 写出配置
//     * @param mergeScopeList   合并范围
//     * @param bgColors         背景色
//     */
//    public static void downExcel(Collection data, List<ExcelColumnConfig> configs, ExcelWriteConfig excelWriteConfig, List<ExcelMergeScope> mergeScopeList, List<ExcelBgColor> bgColors) {
//        //定义目标文件
//        File destFile = TmpFileUtil.getTmpFile(excelWriteConfig.getFileName());
//        File templateFile = new File(excelWriteConfig.getTemplatePath());
//        downExcel(data, configs, destFile, templateFile, excelWriteConfig.isOnlyAlias(), excelWriteConfig.isWriteKeyAsHead(), excelWriteConfig.getStartRowIndex(), mergeScopeList, bgColors);
//    }
//
//
//    /**
//     * 导出excel
//     *
//     * @param data             导出列表
//     * @param configs          配置
//     * @param destFile         目标文件
//     * @param templateFile     模板文件
//     * @param onlyAlias        是否仅写出有别名的列
//     * @param isWriteKeyAsHead 是否写出标题行
//     * @param mergeScopeList   合并范围
//     * @param bgColors         背景色
//     */
//    @SneakyThrows
//    private static void downExcel(Collection data, List<ExcelColumnConfig> configs, File destFile, File templateFile,
//                                  boolean onlyAlias, boolean isWriteKeyAsHead, int startRowIndex, List<ExcelMergeScope> mergeScopeList, List<ExcelBgColor> bgColors) {
//
//        //写Excel文档
//        BigExcelWriter writer = writerExcel(data, configs, destFile, templateFile, onlyAlias, isWriteKeyAsHead, startRowIndex, mergeScopeList, bgColors);
//
//        //设置响应
//        HttpServletResponse response = ServletStaticUtil.getResponse();
//        response.setContentType("application/vnd.ms-excel;charset=utf-8");
//        response.setHeader("Content-Disposition", "attachment;filename=" + URLEncodeUtil.encode(destFile.getName()));
//
//        //将Excel文档写入输出流
//        ServletOutputStream out = response.getOutputStream();
//        writer.flush(out, true);
//        writer.close();
//
//        //删除临时文件
//        if (templateFile != null) {
//            FileUtil.del(templateFile);
//        }
//        FileUtil.del(destFile);
//    }
//
//
//    /**
//     * 生成Excel文件,返回目标文件
//     *
//     * @param data           数据
//     * @param configs        列配置
//     * @param fileName       文件名
//     * @param mergeScopeList 合并范围
//     * @param bgColors       背景色
//     * @return 目标文件
//     */
//    public static File generateExcel(Collection data, List<ExcelColumnConfig> configs, String fileName, List<ExcelMergeScope> mergeScopeList, List<ExcelBgColor> bgColors) {
//        //定义目标文件
//        File destFile = TmpFileUtil.getTmpFile(fileName);
//
//        //写Excel文档
//        BigExcelWriter writer = writerExcel(data, configs, destFile, null, true, true, 0, mergeScopeList, bgColors);
//
//        //先写出目标文件后给关闭工作簿
//        writer.close();
//
//        //返回目标文件
//        return destFile;
//    }
//
//    /**
//     * 生成Excel文件,返回目标文件
//     *
//     * @param data             数据
//     * @param configs          列配置
//     * @param excelWriteConfig 写出配置
//     * @param mergeScopeList   合并范围
//     * @param bgColors         背景色
//     * @return 目标文件
//     */
//    public static File generateExcel(Collection data, List<ExcelColumnConfig> configs, ExcelWriteConfig excelWriteConfig, List<ExcelMergeScope> mergeScopeList, List<ExcelBgColor> bgColors) {
//        //定义目标文件
//        File destFile = TmpFileUtil.getTmpFile(excelWriteConfig.getFileName());
//
//        //写Excel文档
//        File templateFile = new File(excelWriteConfig.getTemplatePath());
//        BigExcelWriter writer = writerExcel(data, configs, destFile, templateFile, excelWriteConfig.isOnlyAlias(), excelWriteConfig.isWriteKeyAsHead(), excelWriteConfig.getStartRowIndex(), mergeScopeList, bgColors);
//
//        //先写出目标文件后给关闭工作簿
//        writer.close();
//
//        //返回目标文件
//        return destFile;
//    }
//
//    /**
//     * 写Excel文件
//     *
//     * @param data             数据
//     * @param configs          列配置
//     * @param destFile         目标文件
//     * @param templateFile     模板文件
//     * @param onlyAlias        是否仅写出有别名的列
//     * @param isWriteKeyAsHead 是否写出标题行
//     * @param startRowIndex    起始行
//     * @param mergeScopeList   单元格合并信息
//     * @param bgColors         背景色
//     * @return excel写对象
//     */
//    private static BigExcelWriter writerExcel(Collection data, List<ExcelColumnConfig> configs, File destFile, File templateFile, boolean onlyAlias,
//                                              boolean isWriteKeyAsHead, int startRowIndex, List<ExcelMergeScope> mergeScopeList, List<ExcelBgColor> bgColors) {
//        BigExcelWriter writer = null;
//        if (templateFile != null) {
//            //根据模板写数据
//            try {
//                SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook(new XSSFWorkbook(templateFile), -1);
//                writer = new BigExcelWriter(sxssfWorkbook, "sheet1");
//            } catch (Exception ex) {
//                log.error("构造BigExcelWriter异常：", ex);
//            }
//            log.info("exportExcel templateFile：" + templateFile);
//        } else {
//            //直接根据目标文件写数据
//            writer = new BigExcelWriter(-1);
//        }
//        //添加列别名
//        if (CollUtil.isNotEmpty(configs)) {
//            for (ExcelColumnConfig config : configs) {
//                writer.addHeaderAlias(config.getOrigin(), config.getAlias());
//            }
//        }
//
//        //设置起始行
//        writer.setCurrentRow(startRowIndex);
//
//        // 只写出加了别名的字段
//        writer.setOnlyAlias(onlyAlias);
//
//        // 合并单元格
//        if (CollUtil.isNotEmpty(mergeScopeList)) {
//            for (ExcelMergeScope mergeScope : mergeScopeList) {
//                writer.merge(mergeScope.getFirstRow(), mergeScope.getLastRow(), mergeScope.getFirstColumn(), mergeScope.getLastColumn(), null, null);
//            }
//        }
//
//        // 写入数据
//        writer.write(data, isWriteKeyAsHead);
//
//        // 调整列样式（在写入数据后方可调整样式）
//        if (CollUtil.isNotEmpty(configs)) {
//            //从标题行下一行开始覆写整列单元格样式
//            if (isWriteKeyAsHead) {
//                startRowIndex++;
//            }
//
//            // 背景色转成Map
//            Map<String, Short> cellBgColorMap = getCellBgColorMap(bgColors);
//
//            //遍历列
//            int columnIndex = 0;
//            for (ExcelColumnConfig config : configs) {
//                //创建新样式
//                CellStyle cellStyle = writer.createCellStyle();
//                //设置边框
//                StyleUtil.setBorder(cellStyle, BorderStyle.THIN, IndexedColors.BLACK);
//                //设置对齐方式
//                if (config.getHorizontalAlignment() != null) {
//                    cellStyle.setAlignment(config.getHorizontalAlignment());
//                }
//                cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
//                //设置数据格式
//                if (config.getFmt() != null) {
//                    cellStyle.setDataFormat(config.getFmt());
//                }
//                // 修改整列格式
//                writer.setColumnStyleIfHasData(columnIndex, startRowIndex, cellStyle);
//
//                // 统一设置百分比数据样式（针对一列中数字+百分比混合的列特殊处理）
//                setPercentDataFmt(startRowIndex, writer, columnIndex, config, cellStyle);
//
//                // 设置背景色
//                setCellBgColor(startRowIndex, writer, cellBgColorMap, columnIndex);
//
//                //列号递增
//                columnIndex++;
//            }
//        }
//
//        // 设置列宽（不论有无模板均支持调整列宽）
//        int columnIndex = 0;
//        for (ExcelColumnConfig config : configs) {
//            if (config.getWidth() != null) {
//                writer.setColumnWidth(columnIndex++, config.getWidth());
//            }
//        }
//
//        //手动设置目标输出文件
//        writer.setDestFile(destFile);
//        log.info("exportExcel destFile：" + destFile);
//        return writer;
//    }
//
//    private static void setPercentDataFmt(int startRowIndex, BigExcelWriter writer, int columnIndex, ExcelColumnConfig config, CellStyle cellStyle) {
//        if (config.getFmt() == null || (config.getFmt() != 9 && config.getFmt() != 0xa)) {
//            CellStyle percentCellStyle = writer.createCellStyle();
//            percentCellStyle.cloneStyleFrom(cellStyle);
//            percentCellStyle.setDataFormat((short) 0xa);
//            for (int r = startRowIndex; r < writer.getRowCount(); r++) {
//                Cell cell = writer.getCell(columnIndex, r);
//                Object cellValue = CellUtil.getCellValue(cell);
//                String valStr = cellValue == null ? StrUtil.EMPTY : String.valueOf(cellValue);
//                valStr = StrUtil.trim(valStr);
//                //百分比字符串 转换成 百分比数字格式
//                if (StrUtil.isNotBlank(valStr) && valStr.contains("%")) {
//                    String valStrWithoutPercent = valStr.replace("%", "");
//                    boolean isPercent = false;
//                    double val = 0.0;
//                    if (NumberUtil.isDouble(valStrWithoutPercent)) {
//                        val = Double.parseDouble(valStrWithoutPercent) / 100;
//                        isPercent = true;
//                    } else if (NumberUtil.isLong(valStrWithoutPercent)) {
//                        val = Long.parseLong(valStrWithoutPercent) / 100.0;
//                        isPercent = true;
//                    }
//                    if (isPercent) {
//                        //重新设置值
//                        cell.setCellValue(val);
//
//                        //重新设置单元格格式
//                        cell.setCellStyle(percentCellStyle);
//                    }
//                }
//            }
//        }
//    }
//
//    private static Map<String, Short> getCellBgColorMap(List<ExcelBgColor> bgColors) {
//        Map<String, Short> cellBgColorMap = new HashMap<>();
//        if (CollUtil.isNotEmpty(bgColors)) {
//            for (ExcelBgColor bgColor : bgColors) {
//                cellBgColorMap.put(bgColor.getCellLocation(), bgColor.getColor());
//            }
//        }
//        return cellBgColorMap;
//    }
//
//    private static void setCellBgColor(int startRowIndex, BigExcelWriter writer, Map<String, Short> cellBgColorMap, int columnIndex) {
//        for (int rowIndex = startRowIndex; rowIndex < writer.getRowCount(); rowIndex++) {
//            String colName = ExcelUtil.indexToColName(columnIndex);
//            String location = colName + (rowIndex + 1);
//            Short color = cellBgColorMap.get(location);
//            if (color != null) {
//                ExcelUtilExt.setCellBgColor(writer, columnIndex, rowIndex, color);
//            }
//        }
//    }
//}