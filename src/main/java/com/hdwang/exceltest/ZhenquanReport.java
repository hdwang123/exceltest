package com.hdwang.exceltest;

import cn.hutool.json.JSONUtil;
import com.hdwang.exceltest.exceldata.ColIndex;

/**
 * 证券月报
 */
public class ZhenquanReport {

    /**
     * 名称
     */
    @ColIndex(name = "B")
    private String name;

    /**
     * 数值
     */
    @ColIndex(index = 2)
    private String value;

    /**
     * 数值
     */
    @ColIndex(name = "E")
    private String value2;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }

    public String getValue2() {
        return value2;
    }

    public void setValue2(String value2) {
        this.value2 = value2;
    }

    @Override
    public String toString() {
        return JSONUtil.toJsonStr(this);
    }
}
