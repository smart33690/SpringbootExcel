package com.xiaopotian.demo;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * 字体处理
 */
public class test08 {
    public static void main(String[] args) throws IOException {
        // 定义一个新的工作簿
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("我的工作簿");
        Row row = sheet.createRow(0);

        // 创建一个字体处理类
        Font font = wb.createFont();
        // 设置字体大小
        font.setFontHeightInPoints((short) 24);
        font.setFontName("微软雅黑");
        // 设置斜体
        font.setItalic(true);
        // 设置中划线
        font.setStrikeout(true);
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setFont(font);

        Cell cell = row.createCell(0);
        cell.setCellValue("春天在哪里");
        cell.setCellStyle(cellStyle);

        FileOutputStream fos;

        try {
            fos = new FileOutputStream("f:\\我的工作簿.xls");
            wb.write(fos);
            wb.close();
            fos.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    }
