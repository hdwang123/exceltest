package com.hdwang.exceltest.model;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 列值截去器
 *
 * @author wanghuidong
 * @date 2022/1/27 16:12
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(value = {ElementType.FIELD})
public @interface ColValTrimmer {

    /**
     * 设置列值最大长度，超出将被截去（默认不截去）
     *
     * @return 列值最大长度
     */
    int maxLength() default -1;

    /**
     * 正则式数组，匹配到的内容将被截去
     * 例如：{"\\d+[\\.、，,:：][ ]*"}
     *
     * @return 正则式数组
     */
    String[] trimRegex() default {};


}
