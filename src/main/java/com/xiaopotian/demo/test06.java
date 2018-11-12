package com.xiaopotian.demo;

import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * 单元格填充色和颜色操作
 */
public class test06 {
    public static void main(String[] args) throws IOException {
        Workbook wb = new HSSFWorkbook(); // 定义一个新的工作簿
        Sheet sheet = wb.createSheet("第一个Sheet页"); // 创建第一个Sheet页
        Row row = sheet.createRow(1); // 创建一个行

        Cell cell = row.createCell(1);
        cell.setCellValue("XX");
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setFillBackgroundColor(IndexedColors.BLUE.getIndex()); // 背景色
        cellStyle.setFillPattern(CellStyle.BIG_SPOTS);
        cell.setCellStyle(cellStyle);

        Cell cell2 = row.createCell(2);
        cell2.setCellValue("YYY");
        CellStyle cellStyle2 = wb.createCellStyle();
        cellStyle2.setFillForegroundColor(IndexedColors.RED.getIndex()); // 前景色
        cellStyle2.setFillPattern(CellStyle.SOLID_FOREGROUND);
        cell2.setCellStyle(cellStyle2);

        FileOutputStream fileOut = new FileOutputStream("f:\\工作簿02.xls");
        wb.write(fileOut);
        wb.close();
        fileOut.close();
    }

}
