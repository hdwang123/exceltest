package com.hdwang.exceltest.validate.validator;

import cn.hutool.core.util.StrUtil;
import cn.hutool.json.JSONUtil;
import com.hdwang.exceltest.model.CellData;
import com.hdwang.exceltest.model.ExcelData;
import com.hdwang.exceltest.validate.ErrorCode;
import com.hdwang.exceltest.validate.ValidateResult;

import java.util.List;

/**
 * IN值校验器
 * 判断单元格的值是否在某些值内
 *
 * @author wanghuidong
 * @date 2022/1/27 16:12
 */
public class InValidator extends AbstractValidator {

    private List<String> valueList;
    private String errorMsg;

    /**
     * 构造器
     *
     * @param valueList 待比较的值列表
     */
    public InValidator(List<String> valueList) {
        this.valueList = valueList;
    }

    /**
     * 构造器
     *
     * @param valueList 待比较的值列表
     * @param errorMsg  错误提示消息
     */
    public InValidator(List<String> valueList, String errorMsg) {
        this.valueList = valueList;
        this.errorMsg = errorMsg;
    }


    @Override
    public void validate(CellData cellData, ExcelData excelData, ValidateResult result) {
        String cellDataValue = cellData.getValue() == null ? StrUtil.EMPTY : String.valueOf(cellData.getValue());
        //数值比较，转换为BigDecimal类型进行比较
        if (!this.valueList.contains(cellDataValue)) {
            result.setErrorCode(ErrorCode.NOT_IN);
            result.setMsg(result.getMsg() + ",期望值：" + JSONUtil.toJsonStr(this.valueList));
            if (StrUtil.isNotEmpty(errorMsg)) {
                result.setMsg(errorMsg);
            }
        }
    }

}
