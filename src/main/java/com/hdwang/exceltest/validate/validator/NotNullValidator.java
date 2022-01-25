package com.hdwang.exceltest.validate.validator;

import cn.hutool.core.util.StrUtil;
import com.hdwang.exceltest.exceldata.CellData;
import com.hdwang.exceltest.exceldata.ExcelData;
import com.hdwang.exceltest.validate.ErrorCode;
import com.hdwang.exceltest.validate.ValidateResult;

/**
 * 非空校验器
 * 判断单元格的值是否为空
 */
public class NotNullValidator extends AbstractValidator {

    @Override
    public void validate(CellData cellData, ExcelData excelData, ValidateResult result) {
        if (cellData.getValue() == null || StrUtil.isBlank(String.valueOf(cellData.getValue()))) {
            result.setErrorCode(ErrorCode.DATA_EMPTY);
        }
    }
}
