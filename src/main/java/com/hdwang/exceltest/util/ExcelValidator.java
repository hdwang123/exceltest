package com.hdwang.exceltest.util;

import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.cell.CellLocation;
import com.hdwang.exceltest.model.CellData;
import com.hdwang.exceltest.model.ExcelData;
import com.hdwang.exceltest.validate.ErrorCode;
import com.hdwang.exceltest.validate.ValidateResult;
import com.hdwang.exceltest.validate.validator.Validator;


import java.util.ArrayList;
import java.util.List;

/**
 * 表格校验工具类
 *
 * @author wanghuidong
 * @date 2022/1/27 16:12
 */
public class ExcelValidator {

    /**
     * 校验行列范围内所有单元格
     *
     * @param excelData     表格数据
     * @param startRowIndex 起始行号（从0开始）
     * @param endRowIndex   结束行号（从0开始,包括此行）
     * @param startColIndex 起始列号（从0开始）
     * @param endColIndex   结束列号（从0开始,包括此列）
     * @param validators    校验器（可以多个）
     * @return 校验结果（仅返回校验失败的）
     */
    public static List<ValidateResult> validate(ExcelData excelData, int startRowIndex, int endRowIndex, int startColIndex, int endColIndex, Validator... validators) {
        List<ValidateResult> results = new ArrayList<>();
        for (int r = startRowIndex; r <= endRowIndex; r++) {
            for (int c = startColIndex; c <= endColIndex; c++) {
                CellData cellData = excelData.getCellData(r, c);
                if (cellData == null) {
                    cellData = new CellData();
                    cellData.setRowIndex(r);
                    cellData.setCellIndex(c);
                }
                for (Validator validator : validators) {
                    ValidateResult result = validator.validate(cellData, excelData);
                    if (ErrorCode.OK != result.getErrorCode()) {
                        results.add(result);
                        //单元格校验遇错终止
                        break;
                    }
                }
            }
        }
        return results;
    }

    /**
     * 校验行列范围内所有单元格
     *
     * @param excelData     表格数据
     * @param startRowIndex 起始行号（从0开始）
     * @param endRowIndex   结束行号（从0开始,包括此行）
     * @param startColName  起始列名（从A开始）
     * @param endColName    结束列名（从A开始,包括此列）
     * @param validators    校验器（可以多个）
     * @return 校验结果（仅返回校验失败的）
     */
    public static List<ValidateResult> validate(ExcelData excelData, int startRowIndex, int endRowIndex, String startColName, String endColName, Validator... validators) {
        int startCellIndex = ExcelUtil.colNameToIndex(startColName + "0");
        int endCellIndex = ExcelUtil.colNameToIndex(endColName + "0");
        return validate(excelData, startRowIndex, endRowIndex, startCellIndex, endCellIndex, validators);
    }

    /**
     * 校验指定单元格
     *
     * @param excelData  表格数据
     * @param rowIndex   行号（从0开始）
     * @param colIndex   列号（从0开始）
     * @param validators 校验器（可以多个）
     * @return 校验结果
     */
    public static ValidateResult validate(ExcelData excelData, int rowIndex, int colIndex, Validator... validators) {
        CellData cellData = excelData.getCellData(rowIndex, colIndex);
        if (cellData == null) {
            cellData = new CellData();
            cellData.setRowIndex(rowIndex);
            cellData.setCellIndex(colIndex);
        }
        for (Validator validator : validators) {
            ValidateResult result = validator.validate(cellData, excelData);
            if (ErrorCode.OK.getCode() != result.getCode()) {
                //单元格校验遇错终止
                return result;
            }
        }
        //没有错，则返回校验成功
        ValidateResult result = new ValidateResult();
        result.setCellData(cellData);
        result.setErrorCode(ErrorCode.OK);
        return result;
    }

    /**
     * 校验指定单元格
     *
     * @param excelData  表格数据
     * @param rowIndex   行号（从0开始）
     * @param colName    列名（从A开始）
     * @param validators 校验器（可以多个）
     * @return 校验结果
     */
    public static ValidateResult validate(ExcelData excelData, int rowIndex, String colName, Validator... validators) {
        int colIndex = ExcelUtil.colNameToIndex(colName + "0");
        return validate(excelData, rowIndex, colIndex, validators);
    }


    /**
     * 校验指定单元格
     *
     * @param excelData   表格数据
     * @param locationRef 单元格位置（例：A1）
     * @param validators  校验器（可以多个）
     * @return 校验结果
     */
    public static ValidateResult validate(ExcelData excelData, String locationRef, Validator... validators) {
        CellLocation cellLocation = ExcelUtil.toLocation(locationRef);
        return validate(excelData, cellLocation.getY(), cellLocation.getX(), validators);
    }
}
