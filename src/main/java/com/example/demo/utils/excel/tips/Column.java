package com.example.demo.utils.excel.tips;

import java.lang.annotation.*;

/**
 * 导出Excel时标注列
 * 包含列明
 * Created by mw4157 on 16/2/17.
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Repeatable(Columns.class)
public @interface Column{

    /** excel列名,就是表头 */
    String value() default "";

    /**
     * 列表排序, 小号在前面, 默认的话就按实体里字段顺序了
     */
    int index() default 0;

    /** 属于哪个@Excel, 填@Excel.value */
    String[] belong() default "";
}
