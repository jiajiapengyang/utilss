package com.example.demo.utils.excel.goods;

import java.lang.reflect.Method;

/**
 * 列信息实体
 */
public class FieldColumn {
    /** 列对应的实体中的getter方法, 如 getName()*/
    private Method getter;
    /**
     * 列对应的实体中的setter方法, 如 setName(name)
     */
    private Method setter;
    /** Excel中的列名,即表头 */
    private String columnName;
    /** 顺序索引 */
    private int index;

    public FieldColumn(Method getter, Method setter, String columnName, int index) {
        this.getter = getter;
        this.setter = setter;
        this.columnName = columnName;
        this.index = index;
    }


    public Method getGetter() {
        return getter;
    }

    public void setGetter(Method getter) {
        this.getter = getter;
    }

    public Method getSetter() {
        return setter;
    }

    public void setSetter(Method setter) {
        this.setter = setter;
    }

    public String getColumnName() {
        return columnName;
    }

    public void setColumnName(String columnName) {
        this.columnName = columnName;
    }

    public int getIndex() {
        return index;
    }

    public void setIndex(int index) {
        this.index = index;
    }

    @Override
    public String toString() {
        return "FieldColumn{" +
                "getter=" + getter +
                ", setter=" + setter +
                ", columnName='" + columnName + '\'' +
                ", index=" + index +
                '}';
    }
}