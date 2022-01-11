package com.hdwang.exceltest.validate;

import cn.hutool.json.JSONUtil;
import com.hdwang.exceltest.exceldata.CellData;

/**
 * 校验结果
 */
public class ValidateResult {

    /**
     * 单元格数据
     */
    private CellData cellData;

    /**
     * 错误代码对象
     */
    private ErrorCode errorCode = ErrorCode.OK;

    /**
     * 错误代码
     */
    private int code = 0;

    /**
     * 错误提示
     */
    private String msg;

    public CellData getCellData() {
        return cellData;
    }

    public void setCellData(CellData cellData) {
        this.cellData = cellData;
    }

    public ErrorCode getErrorCode() {
        return errorCode;
    }

    public void setErrorCode(ErrorCode errorCode) {
        this.errorCode = errorCode;
        this.code = errorCode.getCode();
        this.msg = errorCode.getDesc();
    }

    public int getCode() {
        return code;
    }

    public void setCode(int code) {
        this.code = code;
    }

    public String getMsg() {
        return msg;
    }

    public void setMsg(String msg) {
        this.msg = msg;
    }

    @Override
    public String toString() {
        return JSONUtil.toJsonStr(this);
    }
}
