package com.hdwang.exceltest.validate.validator;


import com.hdwang.exceltest.model.CellData;
import com.hdwang.exceltest.model.ExcelData;
import com.hdwang.exceltest.validate.ErrorCode;
import com.hdwang.exceltest.validate.ValidateResult;
import lombok.extern.slf4j.Slf4j;

/**
 * 校验器抽象类
 *
 * @author wanghuidong
 * @date 2022/1/27 16:12
 */
@Slf4j
public abstract class AbstractValidator implements Validator {

    @Override
    public ValidateResult validate(CellData cellData, ExcelData excelData) {
        ValidateResult result = new ValidateResult();
        result.setCellData(cellData);
        try {
            this.validate(cellData, excelData, result);
        } catch (Exception ex) {
            log.error("校验失败：" + ex.getMessage(), ex);
            result.setErrorCode(ErrorCode.VALIDATE_FAILED);
            result.setMsg(result.getMsg() + ",请检查数据格式是否有误");
        }
        return result;
    }

    /**
     * 校验单元格数据
     *
     * @param cellData  单元格数据
     * @param excelData 表格数据
     * @param result    校验结果
     */
    public abstract void validate(CellData cellData, ExcelData excelData, ValidateResult result);
}
