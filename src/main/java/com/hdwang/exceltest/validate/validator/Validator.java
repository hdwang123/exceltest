package com.hdwang.exceltest.validate.validator;

import com.hdwang.exceltest.exceldata.CellData;
import com.hdwang.exceltest.exceldata.ExcelData;
import com.hdwang.exceltest.validate.ValidateResult;

/**
 * 校验器
 */
public interface Validator {

    /**
     * 校验单元格数据
     *
     * @param cellData  单元格数据
     * @param excelData 表格数据
     * @return 校验结果
     */
    ValidateResult validate(CellData cellData, ExcelData excelData);
}
