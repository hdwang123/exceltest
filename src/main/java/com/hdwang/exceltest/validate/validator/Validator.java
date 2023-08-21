package com.hdwang.exceltest.validate.validator;


import com.hdwang.exceltest.model.CellData;
import com.hdwang.exceltest.model.ExcelData;
import com.hdwang.exceltest.validate.ValidateResult;

/**
 * 校验器接口
 *
 * @author wanghuidong
 * @date 2022/1/27 16:12
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
