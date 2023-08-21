package com.hdwang.exceltest.model;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 列号,用于给bean对象的属性赋指定列的值
 *
 * @author wanghuidong
 * @date 2022/1/27 16:12
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(value = {ElementType.FIELD})
public @interface ColIndex {

    /**
     * 列索引号（从0开始），与name二者填一个即可，优先级高于name
     *
     * @return
     */
    int index() default -1;

    /**
     * 列名称(从A开始)，与index二者填一个即可，优先级低于index
     *
     * @return
     */
    String name() default "";


}
