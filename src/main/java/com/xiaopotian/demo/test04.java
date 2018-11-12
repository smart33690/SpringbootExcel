package com.xiaopotian.demo;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * 单元格对齐方式
 */
public class test04 {
    private static CellStyle cellStyle;
    public static void main(String[] args) throws IOException {
        Workbook wb = new HSSFWorkbook();// 定义一个新的工作簿
        Sheet sheet = wb.createSheet("三年级（2）班花名册");
        Row row = sheet.createRow(0);

        // 设置行的高度
        row.setHeightInPoints(30);
        Cell cell1 = createCellWithMyStyle(wb,row,(short)0,HSSFCellStyle.ALIGN_LEFT,HSSFCellStyle.VERTICAL_CENTER);
        cell1.setCellValue(new HSSFRichTextString("在希望的田野上"));
        Cell cell2 = createCellWithMyStyle(wb,row,(short)1,HSSFCellStyle.ALIGN_RIGHT,HSSFCellStyle.VERTICAL_TOP);
        cell2.setCellValue(new HSSFRichTextString("春天的故事"));
        FileOutputStream fileOut=new FileOutputStream("f:\\工作簿.xls");
        wb.write(fileOut);
        fileOut.close();
    }

    /** * 创建一个单元格 ，并且这个单元格使用了我们自己设置的样式
     * * @param wb Workbook 对象
     * * @param row Row 对象
     * * @param column
     * 设置单元格在 1 行中的第几列
     * * @param align 水平对齐方式
     * * @param vertical 垂直对齐方式
     * * @return
     * */
    private static Cell createCellWithMyStyle(Workbook wb,Row row,short column,short align,short vertical){
        Cell cell = row.createCell(column);
        cellStyle =wb.createCellStyle();
        // 设置单元格水平方向对其方式（默认左对齐）
        cellStyle.setAlignment(align);
        // 设置单元格垂直方向对其方式
        cellStyle.setVerticalAlignment(vertical);
        // 设置单元格样式
        cell.setCellStyle(cellStyle);
        return cell;
    }

}
