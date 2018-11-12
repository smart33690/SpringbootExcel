package com.xiaopotian.demo;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileOutputStream;
import java.io.IOException;

public class sunacTest {
    public static void main(String[] args) throws IOException {
        // 定义一个新的工作簿
        Workbook wb = new HSSFWorkbook();
        Sheet sheet=wb.createSheet("π计划项目统计（总）");
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        // 创建一个字体处理类
        Font font = wb.createFont();
        // 设置字体大小
        font.setFontHeightInPoints((short) 18);
        font.setFontName("微软雅黑");
        font.setBold(true);
        CellStyle cellStyle = wb.createCellStyle();
        //设置字体，上边已经创建了
        cellStyle.setFont(font);
        //设置中央对齐
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cell.setCellValue("π计划精装修项目录入统计表");
        cell.setCellStyle(cellStyle);
        row.setHeightInPoints(sheet.getDefaultRowHeightInPoints() * 2);
        // 创建合并单元格 应该是 Sheet 的任务
        // 合并了 第 2 行至第 6 行，第 4 列至第 9 列的单元格
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 21));
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 6));
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 7, 15));
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 19, 21));

        Font font1 = wb.createFont();
        font1.setFontHeightInPoints((short) 13);
        font1.setFontName("微软雅黑");
        font1.setBold(true);
        CellStyle cellStyle1 = wb.createCellStyle();
        //设置字体，上边已经创建了
        cellStyle1.setFont(font1);
        Row row1 = sheet.createRow(1);
        Cell cell1 = row1.createCell(0);
        cell1.setCellValue("方案总数：");
        cell1.setCellStyle(cellStyle1);
        Cell cell2 = row1.createCell(7);
        cell2.setCellValue("58个");
        cell2.setCellStyle(cellStyle1);
        Cell cell3 = row1.createCell(16);
        cell3.setCellValue("日期：");
        cell3.setCellStyle(cellStyle1);
        Cell cell4 = row1.createCell(17);
        cell4.setCellValue("2018/11/03");
        cell4.setCellStyle(cellStyle1);
        Cell cell5 = row1.createCell(18);
        cell5.setCellValue("统计周期：");
        cell5.setCellStyle(cellStyle1);
        Cell cell6 = row1.createCell(19);
        cell6.setCellValue("2017/07/02至2018/07/02");
        cell6.setCellStyle(cellStyle1);

        FileOutputStream fos = new FileOutputStream(
                "f:\\sunac工作簿.xls");
        // 使用工作簿提供的 write 方法向文件输出流输出
        wb.write(fos);
        fos.close();
        wb.close();
    }
}
