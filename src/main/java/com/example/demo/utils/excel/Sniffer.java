package com.example.demo.utils.excel;



import com.example.demo.utils.excel.goods.ExcelException;
import com.example.demo.utils.excel.goods.FieldColumn;
import com.example.demo.utils.excel.tips.Column;
import com.example.demo.utils.excel.tips.Excel;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;

/**
 * 嗅探器
 * 方便发现实体上使用的注解信息
 * Created by mw4157 on 16/6/22.
 */
public class Sniffer {

    private Logger logger = LoggerFactory.getLogger(Sniffer.class);

    /**
     * 找到将要导出的@Excel注解
     * @param <T>           实体泛型
     * @param modelClass    实体类信息
     * @param fileName      将要导出的文件名, null表示只有一个@Excel,而我要的就是他
     * @return              @Excel信息
     */
    public <T> Excel findExcel(Class<T> modelClass, String fileName) {
        Excel excelAnnotation = modelClass.getDeclaredAnnotation(Excel.class);

        // 合法性判断
        if (null == excelAnnotation) {
            throw new ExcelException("未知的导出实体,请先用@Excel注册");
        }

        if (null == fileName && excelAnnotation.value().length > 1) {
            throw new ExcelException("过多的@Excel.value注册, 请选择本次导出的文件名");
        }

        boolean isFix = false;
        if (null != fileName) {
            for (String excelFileName : excelAnnotation.value()) {
                if (fileName.equals(excelFileName)) {
                    isFix = true;
                    break;
                }
            }
        } else {
            isFix = true;
        }

        // 空值判断
        if (!isFix) {
            throw new ExcelException("未找到相匹配的@Excel,输入的文件名为:" + fileName);
        }

        return excelAnnotation;
    }

    /**
     * 找到类中需要导出的字段信息
     * @param <T>           实体泛型
     * @param modelClass    实体类信息
     * @return              导出的列 列表
     */
    public <T> List<FieldColumn> findFieldColumns(Class<T> modelClass, String fileName) {
        List<FieldColumn> fieldColumnList = new ArrayList<>();

        Field[] fields = modelClass.getDeclaredFields();
        for (Field field : fields) {
            Column column = findColumn(field, fileName);
            if (null != column) {
                FieldColumn fieldColumn = createFieldColumn(modelClass, field, column);
                if (null != fieldColumn) {
                    fieldColumnList.add(fieldColumn);
                }
            }
        }
        // 根据index排序
        fieldColumnList.sort((f1, f2) -> f1.getIndex() - f2.getIndex());

        // 空值判断
        if (0 == fieldColumnList.size()) {
            throw new ExcelException("未找到相匹配的@Column,输入的文件名为:" + fileName);
        }
        return fieldColumnList;
    }

    /**
     * 在单个字段中
     * 找到@Column相匹配的字段
     * @param field         字段信息
     * @param fileName      属于哪个导出文件, null表示不要判断了,有@Column就算
     * @return              相匹配的@Column信息, 如果null表示无匹配
     */
    private Column findColumn(Field field, String fileName) {
        Column[] columnAnnotations = field.getDeclaredAnnotationsByType(Column.class);

        if (columnAnnotations.length == 0) {
            return null;
        }
        if (null == fileName && columnAnnotations.length > 1) {
            throw new ExcelException("过多的@Column注册, 请选择本次要导出的文件名");
        }

        Column column;
        if (null == fileName) {
            column = columnAnnotations[0];
        } else {
            // 备选@Column列表
            List<Column> fixColumnList = new ArrayList<>(columnAnnotations.length);

            for (Column columnAnnotation : columnAnnotations) {
                String[] belongFiles = columnAnnotation.belong();
                if (belongFiles.length == 1 && "".equals(belongFiles[0])) {
                    fixColumnList.add(columnAnnotation);
                } else {
                    for(String belongFile : belongFiles) {
                        if (fileName.equals(belongFile)) {
                            fixColumnList.add(columnAnnotation);
                            break;
                        }
                    }
                }
            }

            // 找到最精确的@Column
            column = findMostFixColumn(fixColumnList);
        }

        return column;
    }

    /**
     * 从备选@Column列表中,找到最匹配的@Column,并返回
     * 匹配原则,越是精确的,优先级越高.
     * 即belong数量少的最高
     * @param fixColumn 备选列表
     * @return          最精确的匹配, 无匹配时返回null
     */
    private Column findMostFixColumn(List<Column> fixColumn) {
        Column column;
        // 从备选列表中获得最精确的@Column
        if (fixColumn.size() == 1) {
            // 只有1个合适, 别犹豫,就是她,我的萌
            column = fixColumn.get(0);
        } else {
            // 多个合适时,找最精确的即belong信息最单一, 还是我的萌
            // 如果没有,得到的结果是null
            column = fixColumn.stream().min(
                    (c1, c2) -> {

                        if (c1.belong().length == 1 && "".equals(c1.belong()[0])) {
                            return 1;
                        } else if (c2.belong().length == 1 && "".equals(c2.belong()[0])) {
                            return -1;
                        } else if (c1.belong().length == c2.belong().length) {
                            throw new ExcelException("无法判断@Column精度,请修改使用方法");
                        } else {
                            return c1.belong().length - c2.belong().length;
                        }
                    }
            ).orElse(null);
        }
        return column;
    }

    /**
     * 创建字段对应的getter方法名
     * @param field     字段信息
     * @return          对应的getter方法名
     */
    private String createMethodName(String methodBegin, Field field) {
        return methodBegin + (char) (field.getName().charAt(0) - 32) + field.getName().substring(1);
    }

    /**
     * setter Method begin
     */
    private String setterName() {
        return "set";
    }

    /**
     * getter Method begin
     *
     * @param field field info
     */
    private String getterName(Field field) {
        String methodBegin = "get";
        if (field.getType() == boolean.class) {
            methodBegin = "is";
        }
        return methodBegin;
    }

    /**
     * 构造导出列信息
     * @param modelClass    类型信息
     * @param field         对应字段
     * @param column        列注解
     * @param <T>           实体泛型
     * @return              导出列的信息
     */
    private <T> FieldColumn createFieldColumn(Class<T> modelClass, Field field, Column column) {
        try {
            Method getterMethod = modelClass.getMethod(createMethodName(getterName(field), field));
            Method setterMethod = modelClass.getMethod(createMethodName(setterName(), field), field.getType());
            return new FieldColumn(
                    getterMethod,
                    setterMethod,
                    "".equals(column.value()) ? field.getName() : column.value(),
                    column.index()
            );
        } catch (NoSuchMethodException e) {
            logger.warn("未找到字段 {} 合法的getter方法, 类型:{}", field.getName(), modelClass.toString());
        }

        return null;
    }
}
