package com.xiaopotian.demo;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * 单元格合并
 */
public class test07 {
    public static void main(String[] args) throws IOException {
        Workbook wb = new HSSFWorkbook(); // 定义一个新的工作簿
        Sheet sheet = wb.createSheet("第一个Sheet页"); // 创建第一个Sheet页

        Row row = sheet.createRow(1);
        Cell cell = row.createCell(3);
        cell.setCellValue("从前有座山，山里有座庙。");
        // 创建合并单元格 应该是 Sheet 的任务
        // 合并了 第 2 行至第 6 行，第 4 列至第 9 列的单元格
        sheet.addMergedRegion(new CellRangeAddress(1, 5, 3, 8));

        FileOutputStream fileOut = new FileOutputStream("f:\\工作簿03.xls");
        wb.write(fileOut);
        wb.close();
        fileOut.close();
    }

}
