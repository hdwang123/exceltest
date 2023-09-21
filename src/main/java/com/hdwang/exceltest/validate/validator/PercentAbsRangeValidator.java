package com.hdwang.exceltest.validate.validator;

import cn.hutool.core.util.StrUtil;
import com.hdwang.exceltest.model.CellData;
import com.hdwang.exceltest.model.ExcelData;
import com.hdwang.exceltest.util.NumberUtil;
import com.hdwang.exceltest.validate.ErrorCode;
import com.hdwang.exceltest.validate.ValidateResult;

/**
 * 百分比绝对值范围校验器
 * 判断单元格的百分比绝对值是否在某个范围内，去除单元格中的%号和-号，然后参与比较
 *
 * @author wsc
 * @date 2022/11/18 13:45
 */
public class PercentAbsRangeValidator extends AbstractValidator {

    private Double minValue;
    private Double maxValue;
    private Boolean minInclusive;
    private Boolean maxInclusive;


    /**
     * 构造器
     *
     * @param minValue 最小值（包含）
     * @param maxValue 最大值（包含）
     */
    public PercentAbsRangeValidator(Double minValue, Double maxValue) {
        this(minValue, maxValue, true, true);
    }


    /**
     * 构造器
     *
     * @param minValue     最小值
     * @param maxValue     最大值
     * @param minInclusive 最小值是否包含
     * @param maxInclusive 最大值是否包含
     */
    public PercentAbsRangeValidator(Double minValue, Double maxValue, Boolean minInclusive, Boolean maxInclusive) {
        this.minValue = minValue;
        this.maxValue = maxValue;
        this.minInclusive = minInclusive;
        this.maxInclusive = maxInclusive;
    }

    @Override
    public void validate(CellData cellData, ExcelData excelData, ValidateResult result) {
        String cellDataValue = cellData.getValue() == null ? StrUtil.EMPTY : String.valueOf(cellData.getValue());
        if (StrUtil.isBlank(cellDataValue)) {
            //空的话就让它无事发生
            return;
        } else {
            //去掉百分号、负号、逗号等
            cellDataValue = cellDataValue.replaceAll("[\\-,%]", "");
        }
        if (!NumberUtil.isNormalNumber(cellDataValue)) {
            //简单格式校验，不对就跳过校验，前提：已经用RegexpValidator做过格式校验了
            return;
        }

        boolean validateOk = true;
        if (this.minValue != null) {
            if (this.minInclusive) {
                if (Double.parseDouble(cellDataValue) < minValue) {
                    validateOk = false;
                }
            } else {
                if (Double.parseDouble(cellDataValue) <= minValue) {
                    validateOk = false;
                }
            }
        }
        if (this.maxValue != null) {
            if (this.maxInclusive) {
                if (Double.parseDouble(cellDataValue) > maxValue) {
                    validateOk = false;
                }
            } else {
                if (Double.parseDouble(cellDataValue) >= maxValue) {
                    validateOk = false;
                }
            }
        }

        //校验失败,设置错误提示
        if (!validateOk) {
            result.setErrorCode(ErrorCode.NOT_IN_RANGE);
            String errMsg = String.format("百分比绝对值的取值范围为:%s%s,%s%s",
                    minInclusive ? "[" : "(",
                    minValue == null ? " " : minValue + "%",
                    maxValue == null ? " " : maxValue + "%",
                    maxInclusive ? "]" : ")");
            result.setMsg(errMsg);
        }
    }

}
