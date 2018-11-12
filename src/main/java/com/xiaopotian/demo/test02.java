package com.xiaopotian.demo;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;
import java.util.Date;

/**
 * 遍历工作簿的行和列并获取单元格内容
 */
public class test02 {
    public static void main(String[] args) throws IOException {
            FileInputStream fis = new FileInputStream("f:\\学生名单.xls");
            POIFSFileSystem fs = new POIFSFileSystem(fis);
            HSSFWorkbook wb = new HSSFWorkbook(fs);
            HSSFSheet sheet = wb.getSheetAt(0);
            if (sheet == null) {
                return;
            }
            // 遍历行
            Row row = null;
            Cell cell = null;
            for (int rowNum = 0; rowNum < sheet.getLastRowNum() + 1; rowNum++) {
                row = sheet.getRow(rowNum);
                if (row == null) {
                    continue;
                }
                // 遍历单元格
                for (int cellNum = 0; cellNum < row.getLastCellNum(); cellNum++) {
                    cell = row.getCell(cellNum);
                    System.out.print(getCellDate(cell) + " ");
                }
                System.out.println();
            }
            wb.close();
        }

        private static String getCellDate (Cell cell){
            String return_string = null;
            switch (cell.getCellType()) {
                case HSSFCell.CELL_TYPE_STRING:
                    return_string = cell.getStringCellValue();
                    break;
                case HSSFCell.CELL_TYPE_NUMERIC:
                    return_string = cell.getNumericCellValue() + "";
                    break;
                case HSSFCell.CELL_TYPE_BOOLEAN:
                    return_string = String.valueOf(cell.getBooleanCellValue());
                default:
                    return_string = "";
                    break;
            }
            return return_string;
        }
}
