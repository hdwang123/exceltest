package com.hdwang.exceltest.validate;

/**
 * 校验错误码
 *
 * @author wanghuidong
 * @date 2022/1/27 16:12
 */
public enum ErrorCode {

    /**
     * 校验OK
     */
    OK(0, "校验成功"),

    /**
     * 数据为空
     */
    DATA_EMPTY(10001, "数据为空"),
    /**
     * 格式错误
     */
    FORMAT_ERROR(10002, "格式错误"),
    /**
     * 数值不等
     */
    NOT_EQUAL(10003, "数值不等"),
    /**
     * 计算错误
     */
    CALCULATION_MISTAKE(10004, "计算错误"),
    /**
     * 不在范围内
     */
    NOT_IN_RANGE(10005, "数值不在范围内"),

    /**
     * 增长幅度异常
     */
    INCREASE_ERROR(10006, "增长幅度异常"),
    /**
     * 值不在指定范围内
     */
    NOT_IN(10007, "值不在指定范围内"),

    /**
     * 校验失败（校验过程抛异常了）
     */
    VALIDATE_FAILED(-1, "校验失败");


    private int code;
    private String desc;

    ErrorCode(int code, String desc) {
        this.code = code;
        this.desc = desc;
    }

    public int getCode() {
        return code;
    }

    public void setCode(int code) {
        this.code = code;
    }

    public String getDesc() {
        return desc;
    }

    public void setDesc(String desc) {
        this.desc = desc;
    }
}
