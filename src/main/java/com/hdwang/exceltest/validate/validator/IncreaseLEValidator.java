package com.hdwang.exceltest.validate.validator;

import cn.hutool.core.util.NumberUtil;
import cn.hutool.core.util.StrUtil;
import com.hdwang.exceltest.model.CellData;
import com.hdwang.exceltest.model.ExcelData;
import com.hdwang.exceltest.validate.ErrorCode;
import com.hdwang.exceltest.validate.ValidateResult;

/**
 * 增长幅度校验器
 * 判断单元格的值相较于另一个值的增长幅度是否小于等于某个值
 *
 * @author wanghuidong
 * 时间： 2023/8/15 17:28
 */
public class IncreaseLEValidator extends AbstractValidator {

    /**
     * 待比较的值
     */
    private String compareVal;

    /**
     * 增幅小于等于指定的值
     */
    private double leValue;

    /**
     * 错误提示
     */
    private String errorMsg;

    /**
     * 是否是增长幅度的绝对值进行比较（是：允许正负增长，即变动幅度， 否：就是正增长）
     */
    private boolean abs;

    /**
     * 增长幅度校验器
     *
     * @param compareVal 待比较的值
     * @param leValue    增幅小于等于指定的值
     */
    public IncreaseLEValidator(String compareVal, double leValue) {
        this(compareVal, leValue, null);
    }

    /**
     * 增长幅度校验器
     *
     * @param compareVal 待比较的值
     * @param leValue    增幅小于等于指定的值
     * @param errorMsg   错误提示
     */
    public IncreaseLEValidator(String compareVal, double leValue, String errorMsg) {
        this(compareVal, leValue, errorMsg, false);
    }

    /**
     * 增长幅度校验器
     *
     * @param compareVal 待比较的值
     * @param leValue    增幅小于等于指定的值
     * @param errorMsg   错误提示
     * @param abs        是否是增长幅度的绝对值进行比较（是：允许正负增长，即变动幅度， 否：就是正增长）
     */
    public IncreaseLEValidator(String compareVal, double leValue, String errorMsg, boolean abs) {
        this.compareVal = compareVal;
        this.leValue = leValue;
        this.errorMsg = errorMsg;
        this.abs = abs;
    }

    @Override
    public void validate(CellData cellData, ExcelData excelData, ValidateResult result) {
        boolean validateOk = false;
        String cellDataValue = cellData.getValue() == null ? StrUtil.EMPTY : String.valueOf(cellData.getValue());
        if (NumberUtil.isNumber(cellDataValue) && NumberUtil.isNumber(this.compareVal)) {
            // 计算增长率
            double increaseRate = NumberUtil.sub(NumberUtil.div(cellDataValue, this.compareVal), 1).doubleValue();
            increaseRate = increaseRate * 100;
            if (abs) {
                increaseRate = Math.abs(increaseRate);
            }
            // 比较增长率
            if (increaseRate <= leValue) {
                validateOk = true;
            }
        } else {
            // 数据格式都不对，直接跳过校验
            return;
        }
        //校验失败，返回错误提示
        if (!validateOk) {
            result.setErrorCode(ErrorCode.INCREASE_ERROR);
            if (StrUtil.isNotEmpty(errorMsg)) {
                result.setMsg(errorMsg);
            }
        }
    }
}
