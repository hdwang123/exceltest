package com.hdwang.exceltest.validate.validator;

import cn.hutool.core.util.StrUtil;
import com.hdwang.exceltest.model.CellData;
import com.hdwang.exceltest.model.ExcelData;
import com.hdwang.exceltest.util.NumberUtil;
import com.hdwang.exceltest.validate.ErrorCode;
import com.hdwang.exceltest.validate.ValidateResult;
/**
 * 范围校验器
 * 判断单元格的值是否在某个范围内
 *
 * @author: P.H
 * @version: 1.8
 */
public class RangeValidator extends AbstractValidator {
    private Double minValue;
    private Double maxValue;
    private Boolean minInclusive;
    private Boolean maxInclusive;
    private String errorMsg;

    /**
     * 构造器
     *
     * @param minValue 最小值（包含）
     * @param maxValue 最大值（包含）
     */
    public RangeValidator(Double minValue, Double maxValue) {
        this(minValue, maxValue, true, true, null);
    }

    /**
     * 构造器
     *
     * @param minValue 最小值（包含）
     * @param maxValue 最大值（包含）
     * @param errorMsg 错误提示
     */
    public RangeValidator(Double minValue, Double maxValue, String errorMsg) {
        this(minValue, maxValue, true, true, errorMsg);
    }

    /**
     * 构造器
     *
     * @param minValue     最小值
     * @param maxValue     最大值
     * @param minInclusive 最小值是否包含
     * @param maxInclusive 最大值是否包含
     */
    public RangeValidator(Double minValue, Double maxValue, Boolean minInclusive, Boolean maxInclusive) {
        this(minValue, maxValue, minInclusive, maxInclusive, null);
    }

    /**
     * 构造器
     *
     * @param minValue     最小值
     * @param maxValue     最大值
     * @param minInclusive 最小值是否包含
     * @param maxInclusive 最大值是否包含
     * @param errorMsg     错误提示
     */
    public RangeValidator(Double minValue, Double maxValue, Boolean minInclusive, Boolean maxInclusive, String errorMsg) {
        this.minValue = minValue;
        this.maxValue = maxValue;
        this.minInclusive = minInclusive;
        this.maxInclusive = maxInclusive;
        this.errorMsg = errorMsg;
    }

    @Override
    public void validate(CellData cellData, ExcelData excelData, ValidateResult result) {
        String cellDataValue = cellData.getValue() == null ? StrUtil.EMPTY : String.valueOf(cellData.getValue());
        if (StrUtil.isBlank(cellDataValue)) {
            //空的话就让它无事发生
            return;
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
            if (StrUtil.isNotBlank(errorMsg)) {
                result.setMsg(errorMsg);
            } else {
                String errMsg = String.format("%s,取值范围:%s%s,%s%s",
                        result.getMsg(),
                        minInclusive ? "[" : "(",
                        minValue == null ? " " : minValue,
                        maxValue == null ? " " : maxValue,
                        maxInclusive ? "]" : ")");
                result.setMsg(errMsg);
            }
        }
    }

}
