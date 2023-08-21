package com.hdwang.exceltest.validate.validator;

import cn.hutool.core.util.StrUtil;
import com.hdwang.exceltest.model.CellData;
import com.hdwang.exceltest.model.ExcelData;
import com.hdwang.exceltest.validate.ErrorCode;
import com.hdwang.exceltest.validate.ValidateResult;


/**
 * 非空校验器
 * 判断单元格的值是否为空
 *
 * @author wanghuidong
 * @date 2022/1/27 16:12
 */
public class NotNullValidator extends AbstractValidator {

    @Override
    public void validate(CellData cellData, ExcelData excelData, ValidateResult result) {
        if (cellData.getValue() == null || StrUtil.isBlank(String.valueOf(cellData.getValue()))) {
            result.setErrorCode(ErrorCode.DATA_EMPTY);
        }
    }
}
