package com.xiaopotian.demo;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;

import java.io.*;

/**
 * 单元格中使用换行
 * 思路：（1）设置单元格文本的时候使用”\n”表示换行。
 * （2）使用 style.setWrapText(true); 设置；
 * （3）换行以后还要设置单元格相应的行高和列宽，具体见代码（这部分我查了一下网上大家的用法，说是对中文的支持不好，当然也有解决方案，见下面的参考资料）。
 * sheet.autoSizeColumn(1);
 * sheet.autoSizeColumn(1, true);
 * Sheet的这个方法实现了列的宽度自动适应，第 1 个参数是列号，从 0 开始，第 2 参数的作用是是否考虑合并单元格。
 */
public class test10 {
    public static void main(String[] args) throws IOException {
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet();
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);

        cell.setCellValue("we are young , I love China , I pround I am Chinese\n my name is Smart");
        // 写 “\n” 还不够，还要通过 CellStyle 设置单元格可以换行
        CellStyle style = wb.createCellStyle();
        style.setWrapText(true);
        cell.setCellStyle(style);

        // 调整行高为默认行高的 2 倍
        row.setHeightInPoints(sheet.getDefaultRowHeightInPoints() * 2);
        // 自适应列宽度的方法，由 Sheet 来设置， 第 1 个参数 是列的序数
        sheet.autoSizeColumn(0, true);

        FileOutputStream fos = new FileOutputStream("f:\\测试工作簿.xls");
        wb.write(fos);
        wb.close();
        fos.close();
    }
}
