package com.hdwang.exceltest.validate.validator;

import cn.hutool.core.util.NumberUtil;
import cn.hutool.core.util.StrUtil;
import com.hdwang.exceltest.model.CellData;
import com.hdwang.exceltest.model.ExcelData;
import com.hdwang.exceltest.validate.ErrorCode;
import com.hdwang.exceltest.validate.ValidateResult;

import java.math.BigDecimal;

/**
 * 增长幅度小于等于校验器
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
     * 负增幅小于等于指定的值
     */
    private Double negativeLeValue;

    /**
     * 增增幅小于等于指定的值
     */
    private Double positiveLeValue;

    /**
     * 增幅是否是率（默认：增幅是差值，true：增幅是增长率）
     */
    private boolean rate;

    /**
     * 错误提示
     */
    private String errorMsg;

    /**
     * 增长幅度校验器
     *
     * @param compareVal      待比较的值
     * @param negativeLeValue 负增幅小于等于指定的值
     * @param positiveLeValue 正增幅小于等于指定的值
     * @param rate            增幅是否是率（默认：增幅是差值，true：增幅是增长率）
     */
    public IncreaseLEValidator(String compareVal, double negativeLeValue, double positiveLeValue, boolean rate) {
        this(compareVal, negativeLeValue, positiveLeValue, rate, null);
    }

    /**
     * 增长幅度校验器
     *
     * @param compareVal      待比较的值
     * @param negativeLeValue 负增幅小于等于指定的值
     * @param positiveLeValue 正增幅小于等于指定的值
     * @param rate            增幅是否是率（默认：增幅是差值，true：增幅是增长率）
     * @param errorMsg        错误提示
     */
    public IncreaseLEValidator(String compareVal, double negativeLeValue, double positiveLeValue, boolean rate, String errorMsg) {
        this.compareVal = compareVal;
        this.negativeLeValue = negativeLeValue;
        this.positiveLeValue = positiveLeValue;
        this.rate = rate;
        this.errorMsg = errorMsg;
    }

    @Override
    public void validate(CellData cellData, ExcelData excelData, ValidateResult result) {
        boolean validateOk = true;
        String cellDataValue = cellData.getValue() == null ? StrUtil.EMPTY : String.valueOf(cellData.getValue());
        if (NumberUtil.isNumber(cellDataValue) && NumberUtil.isNumber(this.compareVal)) {
            // 计算增长差值和增长率
            BigDecimal increase = NumberUtil.sub(cellDataValue, this.compareVal);
            double increaseRate = NumberUtil.div(increase, new BigDecimal(this.compareVal)).doubleValue();
            increaseRate = increaseRate * 100;
            if (rate) {
                // 比较增长率
                if (increaseRate < 0) {
                    if (negativeLeValue != null && Math.abs(increaseRate) > Math.abs(negativeLeValue)) {
                        validateOk = false;
                    }
                } else {
                    if (positiveLeValue != null && increaseRate > positiveLeValue) {
                        validateOk = false;
                    }
                }
            } else {
                // 比较增长差值
                double increaseVal = increase.doubleValue();
                if (increaseVal < 0) {
                    if (negativeLeValue != null && Math.abs(increaseVal) > Math.abs(negativeLeValue)) {
                        validateOk = false;
                    }
                } else {
                    if (positiveLeValue != null && increaseVal > positiveLeValue) {
                        validateOk = false;
                    }
                }
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
