package com.xiaopotian.demo;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * 创建用户自定义数据格式
 */
public class test11 {
    public static void main(String[] args) throws IOException {
        Workbook wb = new HSSFWorkbook(); // 定义一个新的工作簿
        Sheet sheet = wb.createSheet("第一个Sheet页"); // 创建第一个Sheet页

        // 方法内部的变量是局部变量，局部变量无须初始化
        CellStyle cellStyle = wb.createCellStyle();
        DataFormat format = wb.createDataFormat();
        Row row;
        Cell cell;

        short rowNum = 0;
        short colNum = 0;


        row = sheet.createRow(rowNum++);
        cell = row.createCell(colNum++);
        cell.setCellValue(123456.25789456);
        // 设置小数点后保留 3 位
        cellStyle.setDataFormat(format.getFormat("0.000"));
        cell.setCellStyle(cellStyle);


        cell = row.createCell(colNum++);
        cell.setCellValue(12345678910.25789456);
        cellStyle = wb.createCellStyle();
        cellStyle.setDataFormat(format.getFormat("#,##0.0000"));
        cell.setCellStyle(cellStyle);

        sheet.autoSizeColumn(0);
        sheet.autoSizeColumn(1);


        FileOutputStream fileOut = new FileOutputStream("f:\\拉拉工作簿.xls");
        wb.write(fileOut);
        wb.close();
        fileOut.close();
    }
}
