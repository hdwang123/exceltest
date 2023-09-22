package com.hdwang.exceltest.util;

import cn.hutool.core.util.StrUtil;

import java.math.BigDecimal;

/**
 * 数字工具类
 *
 * @author wanghuidong
 * 时间： 2022/11/24 16:09
 */
public class NumberUtil extends cn.hutool.core.util.NumberUtil {

    /**
     * 数字转百分比
     *
     * @param decimal 数字
     * @return 百分比字符串
     */
    public static String decimalToPercent(double decimal) {
        return decimalToPercent(String.valueOf(decimal));
    }

    /**
     * 数字转百分比
     *
     * @param decimalStr 数字字符串
     * @return 百分比字符串
     */
    public static String decimalToPercent(String decimalStr) {
        if (cn.hutool.core.util.NumberUtil.isNumber(decimalStr)) {
            return cn.hutool.core.util.NumberUtil.mul(decimalStr, "100").setScale(2, BigDecimal.ROUND_HALF_UP) + "%";
        } else {
            return decimalStr;
        }
    }

    /**
     * 是否是常规数字（整型、小数）
     *
     * @param str 数字字符串
     * @return 是否是常规数字
     */
    public static boolean isNormalNumber(String str) {
        if (StrUtil.isBlank(str)) {
            return false;
        }
        String reg = "^-?[0-9]+(\\.[0-9]+)?$";
        return str.matches(reg);
    }
}
