package com.hdwang.exceltest.validate.validator;

import cn.hutool.core.util.NumberUtil;
import cn.hutool.core.util.StrUtil;
import com.hdwang.exceltest.model.CellData;
import com.hdwang.exceltest.model.ExcelData;
import com.hdwang.exceltest.validate.ErrorCode;
import com.hdwang.exceltest.validate.ValidateResult;

import java.math.BigDecimal;
import java.util.function.Supplier;

/**
 * 大致等值校验器
 * 支持数字、百分比的大致比较，对数字进行保留两位小数后进行比较,支持浮动范围的比较
 */
public class AlmostEqualValidator extends AbstractValidator {
    /**
     * 待比较的值
     */
    private String value;
    /**
     * 错误消息
     */
    private String errorMsg;
    /**
     * 待比较的值的浮动范围
     */
    private double floatRange = 0;

    /**
     * 构造器
     *
     * @param value 待比较的值
     */
    public AlmostEqualValidator(String value) {
        this.value = value;
    }

    /**
     * 构造器
     *
     * @param value    待比较的值
     * @param errorMsg 错误提示消息
     */
    public AlmostEqualValidator(String value, String errorMsg) {
        this.value = value;
        this.errorMsg = errorMsg;
    }

    /**
     * 构造器
     *
     * @param value      待比较的值
     * @param errorMsg   错误提示消息
     * @param floatRange 值浮动范围
     */
    public AlmostEqualValidator(String value, String errorMsg, double floatRange) {
        this.value = value;
        this.errorMsg = errorMsg;
        this.floatRange = floatRange;
    }

    /**
     * 构造器
     *
     * @param valueSupplier 待比较值提供者
     * @param errorMsg      错误提示消息
     */
    public AlmostEqualValidator(Supplier<String> valueSupplier, String errorMsg) {
        this.value = valueSupplier.get();
        this.errorMsg = errorMsg;
    }

    @Override
    public void validate(CellData cellData, ExcelData excelData, ValidateResult result) {
        boolean validateOk = true;

        //处理单元格的值
        String cellDataValue = cellData.getValue() == null ? StrUtil.EMPTY : String.valueOf(cellData.getValue());
        //如果是百分比，则去掉百分比
        cellDataValue = cellDataValue.replaceAll("%", "");
        if (NumberUtil.isNumber(cellDataValue)) {
            //保留两位小数
            cellDataValue = NumberUtil.round(cellDataValue, 2).toString();
            this.value = NumberUtil.round(this.value, 2).toString();

            if (floatRange == 0) {
                // 等值比较：利用BigDecimal字符串比较
                if (!cellDataValue.equals(this.value)) {
                    validateOk = false;
                }
            } else {
                // 范围比较，上下浮动 floatRange
                BigDecimal maxVal = NumberUtil.add(this.value, String.valueOf(this.floatRange));
                BigDecimal minVal = NumberUtil.sub(this.value, String.valueOf(this.floatRange));
                BigDecimal cellVal = new BigDecimal(cellDataValue);
                if (!(cellVal.compareTo(minVal) >= 0 && cellVal.compareTo(maxVal) <= 0)) {
                    validateOk = false;
                }
            }
        } else {
            // 数据格式都不对，直接跳过校验
            return;
        }
        //校验失败，返回错误提示
        if (!validateOk) {
            result.setErrorCode(ErrorCode.NOT_EQUAL);
            result.setMsg(result.getMsg() + ",期望值：" + this.value);
            if (StrUtil.isNotEmpty(errorMsg)) {
                result.setMsg(errorMsg);
            }
        }
    }
}
