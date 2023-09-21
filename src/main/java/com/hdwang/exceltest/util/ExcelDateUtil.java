package com.hdwang.exceltest.util;

import java.util.HashMap;
import java.util.Map;

/**
 * excel日期工具类
 *
 * @author wanghuidong
 * 时间： 2023/9/7 17:28
 */
public class ExcelDateUtil {
    private static final Map<String, String> DATE_FORMAT_MAP = new HashMap<>();
    private static final Map<String, String> TIME_FORMAT_MAP = new HashMap<>();
    private static final Map<String, String> DATE_TIME_FORMAT_MAP = new HashMap<>();

    static {
        // 日期格式映射
        DATE_FORMAT_MAP.put("yyyy\\-mm\\-dd", "yyyy-MM-dd");
        DATE_FORMAT_MAP.put("yyyy/mm/dd", "yyyy/MM/dd");
        DATE_FORMAT_MAP.put("m/d/yy", "yyyy/M/d");
        DATE_FORMAT_MAP.put("yyyy\"年\"m\"月\"d\"日\";@", "yyyy年M月d日");
        DATE_FORMAT_MAP.put("yyyy\"年\"mm\"月\"dd\"日\"", "yyyy年MM月dd日");


        // 时间格式映射
        TIME_FORMAT_MAP.put("hh:mm:ss", "HH:mm:ss");

        // 日期时间格式映射
        DATE_TIME_FORMAT_MAP.put("yyyy\\-mm\\-dd hh:mm:ss", "yyyy-MM-dd HH:mm:ss");
        DATE_TIME_FORMAT_MAP.put("yyyy/mm/dd hh:mm:ss", "yyyy/MM/dd HH:mm:ss");
        DATE_TIME_FORMAT_MAP.put("m/d/yy hh:mm:ss", "yyyy/M/d HH:mm:ss");
    }

    /**
     * 根据excel日期格式转换为java日期格式
     *
     * @param excelDateFormat excel日期格式
     * @return java日期格式
     */
    public static String getJavaDateFormat(String excelDateFormat) {
        if (DATE_FORMAT_MAP.containsKey(excelDateFormat.toLowerCase())) {
            return DATE_FORMAT_MAP.get(excelDateFormat.toLowerCase());
        } else if (TIME_FORMAT_MAP.containsKey(excelDateFormat.toLowerCase())) {
            return TIME_FORMAT_MAP.get(excelDateFormat.toLowerCase());
        } else if (DATE_TIME_FORMAT_MAP.containsKey(excelDateFormat.toLowerCase())) {
            return DATE_TIME_FORMAT_MAP.get(excelDateFormat.toLowerCase());
        } else {
            // 默认格式
            return "yyyy-MM-dd";
        }
    }
}
