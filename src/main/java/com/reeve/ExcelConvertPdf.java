package com.reeve;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.InputStream;
import java.io.OutputStream;

/**
 * excel转换成pdf接口
 * 基于easyExcel和ItextPdf实现
 * @Author Reeve
 * @Date 2021/7/29 10:17
 */
public interface ExcelConvertPdf {

    void convert(String excelPath,String pdfPath);

    void convert(String excelPath, OutputStream outputStream);

    void convert(InputStream inputStream, OutputStream outputStream);

    void convert(HSSFWorkbook workbook, OutputStream outputStream);
}
