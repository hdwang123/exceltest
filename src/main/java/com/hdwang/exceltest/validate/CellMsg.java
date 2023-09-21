package com.hdwang.exceltest.validate;

import cn.hutool.json.JSONUtil;
import lombok.Data;

/**
 * 单元格错误提示
 *
 * @author wanghuidong
 * @date 2022/1/27 16:12
 */
@Data
public class CellMsg {

    /**
     * 行号
     */
    private int rowIndex;

    /**
     * 列号
     */
    private int cellIndex;

    /**
     * 单元格位置,如：A1
     */
    private String location;

    /**
     * 错误提示
     */
    private String msg;

    @Override
    public String toString() {
        return JSONUtil.toJsonStr(this);
    }
}
