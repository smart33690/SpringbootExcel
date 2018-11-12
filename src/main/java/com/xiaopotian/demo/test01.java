package com.xiaopotian.demo;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;
import java.util.Date;

/**
 * 创建一个时间格式的单元格
 */
public class test01 {
    public static void main(String[] args) throws IOException {
        // 定义一个新的工作簿
        Workbook wb = new HSSFWorkbook();
        Sheet sheet1=wb.createSheet("一年级一班");
        // CreationHelper 可以理解为一个工具类，由这个工具类可以获得 日期格式化的一个实例
        CreationHelper createHelper=wb.getCreationHelper();
        // CellStyle 为单元格创建样式的一个接口
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("yyyy-mm-dd"));

        Row row = sheet1.createRow(0);
        // 设置单元格的值为日期类型，这里就涉及到了日期类型值的格式化问题
        Cell c1 = row.createCell(0);
        c1.setCellValue(new Date());
        c1.setCellStyle(cellStyle);

        // 还可以设置单元格的值为 Calendar 的实例
        // Calendar.getInstance();
        // 获取当天指定点上的时间
        Cell c2 = row.createCell(1);
        c2.setCellValue(Calendar.getInstance());
        c2.setCellStyle(cellStyle);

        Sheet sheet2 = wb.createSheet("三年级（1）班学生名单");
        Row row2 = sheet2.createRow(0);
        row2.createCell(0).setCellValue(1);
        row2.createCell(1).setCellValue("一个字符串");
        row2.createCell(2).setCellValue(true);
        row2.createCell(3).setCellValue(HSSFCell.CELL_TYPE_NUMERIC);
        row2.createCell(4).setCellValue(false);


        FileOutputStream fos = new FileOutputStream(
                "f:\\POI工作簿.xls");
        // 使用工作簿提供的 write 方法向文件输出流输出
        wb.write(fos);
        fos.close();
        wb.close();
    }
}
