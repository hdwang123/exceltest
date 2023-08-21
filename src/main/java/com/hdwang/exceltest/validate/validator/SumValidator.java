package com.hdwang.exceltest.validate.validator;

import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.cell.CellLocation;
import com.hdwang.exceltest.model.CellData;
import com.hdwang.exceltest.model.ExcelData;
import com.hdwang.exceltest.validate.ErrorCode;
import com.hdwang.exceltest.validate.ValidateResult;

import java.math.BigDecimal;

/**
 * 求和校验器
 * 判断单元格的和值计算是否准确
 *
 * @author wanghuidong
 * @date 2022/1/27 16:12
 */
public class SumValidator extends AbstractValidator {

    private String startLocationRef;
    private String endLocationRef;

    /**
     * 待计算和值的所有单元格位置
     *
     * @param startLocationRef 起始单元格位置(起始与结束位置可以互换)
     * @param endLocationRef   结束单元格位置(起始与结束位置可以互换)
     */
    public SumValidator(String startLocationRef, String endLocationRef) {
        this.startLocationRef = startLocationRef;
        this.endLocationRef = endLocationRef;
    }

    @Override
    public void validate(CellData cellData, ExcelData excelData, ValidateResult result) {
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
                    BigDecimal value = new BigDecimal(valueStr);
                    sum = sum.add(value);
                }
            }
        }
        //单元格数值全部转成字符串，然后再转换成BigDecimal类型与和值作比较
        String valueStr = String.valueOf(cellData.getValue() == null ? "0" : cellData.getValue());
        BigDecimal cellDataValue = new BigDecimal(valueStr);
        if (cellDataValue.compareTo(sum) != 0) {
            result.setErrorCode(ErrorCode.CALCULATION_MISTAKE);
            result.setMsg(result.getMsg() + ",实际计算结果：" + sum.toString());
        }
    }
}
