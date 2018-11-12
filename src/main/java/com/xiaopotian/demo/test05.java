package com.xiaopotian.demo;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * 单元格边框处理
 */
public class test05 {
    public static void main(String[] args) throws IOException {
        Workbook wb = new HSSFWorkbook();// 定义一个新的工作簿
        Sheet sheet = wb.createSheet("三年级（2）班花名册");
        Row row = sheet.createRow(1);

        // 设置行的高度
        row.setHeightInPoints(30);
        Cell cell = row.createCell(1);
        cell.setCellValue(new HSSFRichTextString("一九七九年，那是一个春天。"));
        CellStyle cellStyle = wb.createCellStyle();
        // 设置底部边框
        cellStyle.setBorderLeft(CellStyle.BORDER_THIN); // 细线
        cellStyle.setLeftBorderColor(IndexedColors.BLUE.getIndex());// 设置颜色

        // 设置右边框
        cellStyle.setBorderRight(CellStyle.BORDER_DASH_DOT);
        cellStyle.setRightBorderColor(IndexedColors.GREEN.getIndex());

        // 设置顶部的边框
        cellStyle.setBorderTop(CellStyle.LEAST_DOTS);
        cellStyle.setTopBorderColor(IndexedColors.BROWN.getIndex());

        cell.setCellStyle(cellStyle);
        FileOutputStream fileOut=new FileOutputStream("f:\\工作簿01.xls");
        wb.write(fileOut);
        wb.close();
        fileOut.close();
    }

}
