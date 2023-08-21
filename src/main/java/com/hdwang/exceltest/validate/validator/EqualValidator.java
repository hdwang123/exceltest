package com.hdwang.exceltest.validate.validator;

import cn.hutool.core.util.StrUtil;
import com.hdwang.exceltest.model.CellData;
import com.hdwang.exceltest.model.ExcelData;
import com.hdwang.exceltest.util.NumberUtil;
import com.hdwang.exceltest.validate.ErrorCode;
import com.hdwang.exceltest.validate.ValidateResult;

import java.math.BigDecimal;
import java.util.function.Supplier;

/**
 * 等值校验器
 * 判断单元格的值是否与某个值相等
 *
 * @author wanghuidong
 * @date 2022/1/27 16:12
 */
public class EqualValidator extends AbstractValidator {

    private String value;
    private String errorMsg;

    /**
     * 构造器
     *
     * @param value 待比较的值
     */
    public EqualValidator(String value) {
        this.value = value;
    }

    /**
     * 构造器
     *
     * @param value    待比较的值
     * @param errorMsg 错误提示消息
     */
    public EqualValidator(String value, String errorMsg) {
        this.value = value;
        this.errorMsg = errorMsg;
    }

    /**
     * 构造器
     *
     * @param valueSupplier 待比较值提供者
     * @param errorMsg      错误提示消息
     */
    public EqualValidator(Supplier<String> valueSupplier, String errorMsg) {
        this.value = valueSupplier.get();
        this.errorMsg = errorMsg;
    }

    @Override
    public void validate(CellData cellData, ExcelData excelData, ValidateResult result) {
        String cellDataValue = cellData.getValue() == null ? StrUtil.EMPTY : String.valueOf(cellData.getValue());
        if (NumberUtil.isNormalNumber(cellDataValue) && NumberUtil.isNormalNumber(this.value)) {
            //数值比较，转换为BigDecimal类型进行比较
            if (new BigDecimal(cellDataValue).compareTo(new BigDecimal(this.value)) != 0) {
                result.setErrorCode(ErrorCode.NOT_EQUAL);
                result.setMsg(result.getMsg() + ",期望值：" + this.value);
                if (StrUtil.isNotEmpty(errorMsg)) {
                    result.setMsg(errorMsg);
                }
            }
        } else {
            //非数值，按照字符串进行比较
            if (!cellDataValue.equals(this.value)) {
                result.setErrorCode(ErrorCode.NOT_EQUAL);
                result.setMsg(result.getMsg() + ",期望值：" + this.value);
                if (StrUtil.isNotEmpty(errorMsg)) {
                    result.setMsg(errorMsg);
                }
            }
        }
    }


}
