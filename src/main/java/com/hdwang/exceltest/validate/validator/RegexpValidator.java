package com.hdwang.exceltest.validate.validator;

import cn.hutool.core.util.StrUtil;
import com.hdwang.exceltest.exceldata.CellData;
import com.hdwang.exceltest.exceldata.ExcelData;
import com.hdwang.exceltest.validate.ErrorCode;
import com.hdwang.exceltest.validate.ValidateResult;

/**
 * 正则式校验器
 * 校验单元格的值的格式是否匹配
 */
public class RegexpValidator implements Validator {

    private String regexp;
    private String errorMsg;

    /**
     * 构造器
     *
     * @param regexp 正则表达式
     */
    public RegexpValidator(String regexp) {
        this.regexp = regexp;
    }

    /**
     * 构造器
     *
     * @param regexp   正则表达式
     * @param errorMsg 错误提示消息
     */
    public RegexpValidator(String regexp, String errorMsg) {
        this.regexp = regexp;
        this.errorMsg = errorMsg;
    }

    @Override
    public ValidateResult validate(CellData cellData, ExcelData excelData) {
        ValidateResult result = new ValidateResult();
        result.setCellData(cellData);
        String cellDataValue = cellData.getValue() == null ? StrUtil.EMPTY : String.valueOf(cellData.getValue());
        if (!cellDataValue.matches(regexp)) {
            result.setErrorCode(ErrorCode.FORMAT_ERROR);
            if (StrUtil.isNotEmpty(this.errorMsg)) {
                result.setMsg(this.errorMsg);
            }
        }
        return result;
    }
}
