package com.example.demo.utils.excel;


import com.example.demo.utils.excel.goods.ExcelException;
import com.example.demo.utils.excel.goods.FieldColumn;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Excel转化器
 * 真正负责Excel处理的类
 * Created by mw4157 on 16/2/22.
 */
public class ExcelConvertor {
    private Logger logger = LoggerFactory.getLogger(ExcelConvertor.class);

    // 时间类型的样式  yyyy-MM-dd HH:mm
    private CellStyle dateCellStyle;

    /**
     * 创建工作簿
     * @param <T>               实体泛型
     * @param data              表数据
     * @param fieldColumnList   列信息
     * @return                  创建好的工作簿
     */
    public <T> SXSSFWorkbook createExcel(List<T> data, List<FieldColumn> fieldColumnList) {

        // 内存中只驻留100行数据
        SXSSFWorkbook workbook = new SXSSFWorkbook(100);

        // 创建日期类型
        dateCellStyle = workbook.createCellStyle();
        dateCellStyle.setDataFormat((short)0x16);

        // 创建工作簿
        Sheet sheet = workbook.createSheet();

        fillHeader(sheet, fieldColumnList);

        fillBody(sheet, fieldColumnList, data);

        return workbook;
    }

    /**
     * 创建动态excel
     * @param headData
     * @param bodyData
     * @return
     */
    public SXSSFWorkbook createDynamicExcel(List headData, List bodyData) {

        // 内存中只驻留100行数据
        SXSSFWorkbook workbook = new SXSSFWorkbook(100);

        // 创建日期类型
        dateCellStyle = workbook.createCellStyle();
        dateCellStyle.setDataFormat((short)0x16);

        // 创建工作簿
        Sheet sheet = workbook.createSheet();

        Row header = sheet.createRow(0);
        fillCell(header,headData);

        for (int j=0; j<bodyData.size(); j++) {
            Row row = sheet.createRow(j+1);
            List<Object> t = (List<Object>) bodyData.get(j);
            fillCell(row,t);
        }

        return workbook;
    }

    /**
     * 填充一行的单元格
     * @param row
     * @param t
     */
    private void fillCell(Row row,List t){
        for (int i=0; i< t.size(); i++) {
            try {
                Object returnValue = t.get(i);
                if (null == returnValue) {
                    row.createCell(i).setCellValue("");
                    continue;
                }
                Class returnType = String.class;
                try{
                    returnType = t.get(i).getClass();
                }catch (Exception e){
                    e.printStackTrace();
                }

                if (returnType == String.class) {
                    row.createCell(i).setCellValue(String.valueOf(returnValue));
                } else if (returnType == int.class || returnType == Integer.class) {
                    row.createCell(i).setCellValue((Integer) returnValue);
                } else if (returnType == short.class || returnType == Short.class) {
                    row.createCell(i).setCellValue((Short) returnValue);
                } else if (returnType == long.class || returnType == Long.class) {
                    row.createCell(i).setCellValue((Long) returnValue);
                } else if (returnType == boolean.class || returnType == Boolean.class) {
                    row.createCell(i).setCellValue((Boolean)returnValue);
                } else if (returnType == double.class || returnType == Double.class) {
                    row.createCell(i).setCellValue((Double)returnValue);
                } else if (returnType == Date.class) {
                    Cell dateCell = row.createCell(i);
                    dateCell.setCellValue((Date)returnValue);
                    dateCell.setCellStyle(dateCellStyle);
                } else {
                    row.createCell(i).setCellValue(String.valueOf(returnValue));
                }
            } catch (Exception e) {
                logger.error("填充单元格出错, index=" + i + ", 数据内容: " + t.get(i));
            }
        }
    }

    /**
     * 读取Excel内容到List中去
     * @param workbook          工作簿
     * @param fieldColumnList   列信息
     * @param modelClass        实体类型信息
     * @param <T>               泛型类型
     * @return 数据列表
     */
    public <T> List<T> readExcel(Workbook workbook, List<FieldColumn> fieldColumnList, Class<T> modelClass) {
        // 目前只支持1页, 再改进
        Sheet sheet = workbook.getSheetAt(0);

        // 判断是否拥有可读内容
        if (sheet.getLastRowNum() < 1) {
            logger.warn("导入的excel没有有用的内容, 请检查");
            return null;
        }

        // 读取第一行,判断列名
        Row firstRow = sheet.getRow(0);
        if (!readHeader(firstRow, fieldColumnList)) {
            throw new ExcelException("导入的Excel没有与之匹配的列名, 请核实");
        }

        return readBody(sheet, fieldColumnList, modelClass);
    }

    /**
     * 读取Excel内容到List中去
     * @param workbook
     * @return
     */
    public List<Map<String,Object>> readDynamicExcel(Workbook workbook) {
        // 目前只支持1页, 再改进
        Sheet sheet = workbook.getSheetAt(0);

        // 判断是否拥有可读内容
        if (sheet.getLastRowNum() < 1) {
            logger.warn("导入的excel没有有用的内容, 请检查");
            return null;
        }

        // 读取第一行,判断列名
        Row firstRow = sheet.getRow(0);
        List<String> head = readHeader(firstRow);

        return readBody(sheet, head);
    }

    /**
     * 组装一个标头的list
     * @param row
     * @return
     */
    private List<String> readHeader(Row row) {
        List<String> head = new ArrayList<>();
        DecimalFormat decimalFormat = new DecimalFormat("#");
        for(int i = 0 ; i < row.getLastCellNum() ; i++){
            try {
                Cell cell = row.getCell(i);

                switch (cell.getCellType()) {
                    case Cell.CELL_TYPE_BOOLEAN:
                        boolean boolValue = cell.getBooleanCellValue();
                        head.add(String.valueOf(boolValue));
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        if (cell.getCellStyle().getDataFormat() != 0) {
                            Date dateValue = cell.getDateCellValue();
                            SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                            head.add(sdf.format(dateValue));
                        } else {
                            double doubleValue = cell.getNumericCellValue();
                            head.add(decimalFormat.format(doubleValue));
                        }
                        break;
                    case Cell.CELL_TYPE_STRING:
                        String strValue = cell.getStringCellValue();
                        head.add(strValue);
                        break;

                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        }

        return head;
    }

    /**
     * 把整个表数据读取到一个List里,每一行数据是一个Map结构,key为表头里的数据
     * @param sheet
     * @param head 表头list
     * @return
     */
    private List<Map<String,Object>> readBody(Sheet sheet, List<String> head) {
        int totalRow = sheet.getLastRowNum();
        List<Map<String,Object>> data = new ArrayList<>(totalRow);

        // 跳过第一行列名
        for (int i = 1; i <= totalRow; i++) {
            Map<String,Object> params = new HashMap<>();
            Row row = sheet.getRow(i);
            if(row == null){
                continue;
            }
            for(int j = 0 ; j < head.size() ; j++){
                params.put(head.get(j),readCell(row.getCell(j)));
            }

            data.add(params);
        }

        return data;
    }

    /**
     * 读取每个单元格数据
     * @param cell
     * @return
     */
    private Object readCell(Cell cell){
        if(cell == null){
           return null;
        }
        try {
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_BOOLEAN:
                    return cell.getBooleanCellValue();
                case Cell.CELL_TYPE_NUMERIC:
                    if (cell.getCellStyle().getDataFormat() != 0) {
                       return cell.getDateCellValue();
                    } else {
                        return cell.getNumericCellValue();
                    }
                case Cell.CELL_TYPE_STRING:
                    return cell.getStringCellValue();
                default:
                    return null;
            }
        } catch (Exception e) {
            logger.error("类型转换出错");
            return null;
        }
    }


    /**
     * 创建表头
     * @param sheet             工作簿中的一页工作表
     * @param fieldColumnList   列信息
     */
    private void fillHeader(Sheet sheet, List<FieldColumn> fieldColumnList) {
        Row header = sheet.createRow(0);
        for (int i=0; i<fieldColumnList.size(); i++) {
            header.createCell(i).setCellValue(fieldColumnList.get(i).getColumnName());
        }
    }

    /**
     * 填充Excel内容
     * @param sheet             工作表
     * @param fieldColumnList   列信息
     * @param data              将要填充的数据
     * @param <T>               实体泛型
     */
    private <T> void fillBody(Sheet sheet, List<FieldColumn> fieldColumnList, List<T> data) {
        for (int i=0; i<data.size(); i++) {
            fillRow(sheet.createRow(i+1), fieldColumnList, data.get(i));
        }
    }

    /**
     * 判断sheet是否拥有合法的head, 就是第一行列名要有
     * @param row               第一行
     * @param fieldColumnList   列信息
     * @return true拥有正确的head, 可以继续读取; false-head不对,拒绝读取
     */
    private boolean readHeader(Row row, List<FieldColumn> fieldColumnList) {
        boolean hasHead = true;
        for (int i = 0; i < fieldColumnList.size(); i++) {
            if (!fieldColumnList.get(i).getColumnName().equals(row.getCell(i).getStringCellValue())) {
                hasHead = false;
                break;
            }
        }
        return hasHead;
    }

    /**
     * 读取sheet里的内容
     *
     * @param sheet           工作表
     * @param fieldColumnList 列信息
     * @param modelClass      数据模型信息
     * @param <T>             实体泛型
     * @return 数据列表
     */
    private <T> List<T> readBody(Sheet sheet, List<FieldColumn> fieldColumnList, Class<T> modelClass) {
        int totalRow = sheet.getLastRowNum();
        List<T> data = new ArrayList<>(totalRow);

        // 跳过第一行列名
        for (int i = 1; i <= totalRow; i++) {

            data.add(readRow(sheet.getRow(i), fieldColumnList, modelClass));
        }

        return data;
    }

    /**
     * 填充一个数据实体为一行Excel
     * @param row               被填充的行对象
     * @param fieldColumnList   列信息
     * @param oneData           将要填充的数据
     * @param <T>               实体泛型
     */
    private <T> void fillRow(Row row, List<FieldColumn> fieldColumnList, T oneData) {
        for (int i=0; i<fieldColumnList.size(); i++) {
            try {
                Object returnValue = fieldColumnList.get(i).getGetter().invoke(oneData);
                if (null == returnValue) {
                    row.createCell(i).setCellValue("");
                    continue;
                }
                Class returnType = fieldColumnList.get(i).getGetter().getReturnType();
                if (returnType == String.class) {
                    row.createCell(i).setCellValue(String.valueOf(returnValue));
                } else if (returnType == int.class || returnType == Integer.class) {
                    row.createCell(i).setCellValue((Integer) returnValue);
                } else if (returnType == short.class || returnType == Short.class) {
                    row.createCell(i).setCellValue((Short) returnValue);
                } else if (returnType == long.class || returnType == Long.class) {
                    row.createCell(i).setCellValue(String.valueOf(returnValue));
                } else if (returnType == boolean.class || returnType == Boolean.class) {
                    row.createCell(i).setCellValue((Boolean)returnValue);
                } else if (returnType == double.class || returnType == Double.class) {
                    row.createCell(i).setCellValue((Double)returnValue);
                } else if (returnType == Date.class) {
                    Cell dateCell = row.createCell(i);
                    dateCell.setCellValue((Date)returnValue);
                    dateCell.setCellStyle(dateCellStyle);
                } else {
                    row.createCell(i).setCellValue(String.valueOf(returnValue));
                }
            } catch (Exception e) {
                logger.error("填充单元格出错, index=" + i + ", 数据内容: " + oneData.toString());
            }
        }
    }

    /**
     * 读取一行Excel数据封装到一个数据实体中
     *
     * @param row             行对象
     * @param fieldColumnList 列参数信息列表
     * @param modelClass      数据实体类型信息
     * @param <T>             数据实体泛型
     * @return 数据对象
     */
    private <T> T readRow(Row row, List<FieldColumn> fieldColumnList, Class<T> modelClass) {
        T rowData;
        try {
            rowData = modelClass.newInstance();
        } catch (Exception e) {
            logger.error("反射生成对象出错, " + e.getMessage());
            return null;
        }

        DecimalFormat decimalFormat = new DecimalFormat("#");

        for (int i = 0; i < fieldColumnList.size(); i++) {
            FieldColumn fieldColumn = fieldColumnList.get(i);
            try {
                Cell cell = row.getCell(i);
                Class returnType = fieldColumnList.get(i).getGetter().getReturnType();

                switch (cell.getCellType()) {
                    case Cell.CELL_TYPE_BOOLEAN:
                        boolean boolValue = cell.getBooleanCellValue();
                        if (returnType == boolean.class || returnType == Boolean.class) {
                            fieldColumn.getSetter().invoke(rowData, boolValue);
                        } else if (returnType == String.class) {
                            fieldColumn.getSetter().invoke(rowData, String.valueOf(boolValue));
                        }
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        if (cell.getCellStyle().getDataFormat() != 0) {
                            Date dateValue = cell.getDateCellValue();
                            if (returnType == Date.class) {
                                fieldColumn.getSetter().invoke(rowData, dateValue);
                            } else if (returnType == String.class) {
                                SimpleDateFormat sdf = new SimpleDateFormat(cell.getCellStyle().getDataFormatString());
                                fieldColumn.getSetter().invoke(rowData, sdf.format(dateValue));
                            }
                        } else {
                            double doubleValue = cell.getNumericCellValue();
                            if (returnType == double.class || returnType == Double.class) {
                                fieldColumn.getSetter().invoke(rowData, doubleValue);
                            } else if (returnType == int.class || returnType == Integer.class) {
                                fieldColumn.getSetter().invoke(rowData, (int) doubleValue);
                            } else if (returnType == short.class || returnType == Short.class) {
                                fieldColumn.getSetter().invoke(rowData, (short) doubleValue);
                            } else if (returnType == long.class || returnType == Long.class) {
                                fieldColumn.getSetter().invoke(rowData, (long) doubleValue);
                            } else if (returnType == String.class) {
                                fieldColumn.getSetter().invoke(rowData, decimalFormat.format(doubleValue));
                            }
                        }
                        break;
                    case Cell.CELL_TYPE_STRING:
                        String strValue = cell.getStringCellValue();
                        if (returnType == String.class) {
                            fieldColumn.getSetter().invoke(rowData, strValue);
                        } else {
                            fieldColumn.getSetter().invoke(rowData, returnType.cast(strValue));
                        }
                        break;
                }
            } catch (Exception e) {
                logger.error("封装实体类型出错, index={}, 列名={}", i, fieldColumn.getColumnName());
            }
        }

        return rowData;
    }

}
