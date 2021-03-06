package com.hdwang.exceltest.validate.validator;

import cn.hutool.core.util.StrUtil;
import com.hdwang.exceltest.exceldata.CellData;
import com.hdwang.exceltest.exceldata.ExcelData;
import com.hdwang.exceltest.validate.ErrorCode;
import com.hdwang.exceltest.validate.ValidateResult;

/**
 * 等值校验器
 * 判断单元格的值是否与某个值相等
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

    @Override
    public void validate(CellData cellData, ExcelData excelData, ValidateResult result) {
        String cellDataValue = cellData.getValue() == null ? StrUtil.EMPTY : String.valueOf(cellData.getValue());
        if (isNumber(cellDataValue)) {
            //数值比较，转换为double类型进行比较
            if (Double.parseDouble(cellDataValue) != Double.parseDouble(this.value)) {
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

    /**
     * 判断字符串是否是数值类型(整型、浮点型)
     *
     * @param str 字符串
     * @return 是否是数值类型
     */
    private static boolean isNumber(String str) {
        String reg = "^-?[0-9]+(.[0-9]+)?$";
        return str.matches(reg);
    }
}
