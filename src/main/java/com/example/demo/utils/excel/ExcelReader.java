package com.example.demo.utils.excel;


import com.lianjia.common.excel.goods.ExcelException;
import com.lianjia.common.excel.goods.FieldColumn;
import com.lianjia.common.excel.tips.Excel;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.List;
import java.util.Map;

/**
 * 读取Excel
 * Created by mw4157 on 16/7/19.
 */
public class ExcelReader {

    private Logger logger = LoggerFactory.getLogger(ExcelReader.class);

    // 转化器, list数据与workbook之间的转化
    private ExcelConvertor excelConvertor;
    // 嗅探器, 发现实体类上的注解
    private Sniffer sniffer;

    // 单身狗
    private ExcelReader() {
        excelConvertor = new ExcelConvertor();
        sniffer = new Sniffer();
    }


    private static class ExcelSingle {
        private static ExcelReader instance = new ExcelReader();
    }

    public static ExcelReader instance() {
        return ExcelSingle.instance;
    }

    /**
     * 从request里读取Excel并转化成list结构
     * <p>
     * 针对只参与一张excel导出的实体,可使用此方法
     * 即实体上只有一个@Excel注解
     *
     * @param file       上传的文件
     * @param modelClass 数据的实体类型信息, 需要使用@Excel
     * @param <T>        实体泛型
     * @return 整理成的数据, list结构
     */
    public <T> List<T> importToList(MultipartFile file, Class<T> modelClass) {
        return importToList(file, modelClass, null);
    }

    /**
     * 从request里读取Excel并转化成list结构
     *
     * @param file       上传的文件
     * @param modelClass 数据的实体类型信息, 需要使用@Excel
     * @param fileName   导入的Excel原始文件名, 即在导出时的文件名, 用于实体对应多张表是的映射
     * @param <T>        实体泛型
     * @return 整理成的数据, list结构
     */
    public <T> List<T> importToList(MultipartFile file, Class<T> modelClass, String fileName) {

        // 找到导出文件的信息@Excel
        Excel excelAnnotation = sniffer.findExcel(modelClass, fileName);

        // 找到导出列的信息 @Column
        List<FieldColumn> fieldColumnList = sniffer.findFieldColumns(modelClass, fileName);

        XSSFWorkbook workbook;
        try {
            workbook = new XSSFWorkbook(file.getInputStream());
        } catch (IOException e) {
            throw new ExcelException("导入文件异常:" + e.getMessage());
        }

        return excelConvertor.readExcel(workbook, fieldColumnList, modelClass);
    }

    /**
     * 返回的是每一行为一个list,内容是每一个单元格的key-value结构的,其中key为第一行表示的头
     * 方便动态的表格
     * @param file
     * @return
     */
    public List<Map<String,Object>> importToList(MultipartFile file) {
        XSSFWorkbook workbook;
        try {
            workbook = new XSSFWorkbook(file.getInputStream());
        } catch (IOException e) {
            throw new ExcelException("导入文件异常:" + e.getMessage());
        }

        return excelConvertor.readDynamicExcel(workbook);
    }
}
