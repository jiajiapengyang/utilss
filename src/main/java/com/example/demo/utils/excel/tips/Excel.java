package com.example.demo.utils.excel.tips;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 需要导出成excel的实体
 * Created by mw4157 on 16/2/17.
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface Excel {

    /** 导出的文件叫什么? */
    String[] value() default "没有名字的导出表";

    /** 最大导出条数 */
    int limit() default 1040000;
}
