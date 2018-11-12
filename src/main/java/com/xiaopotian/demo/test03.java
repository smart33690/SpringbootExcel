package com.xiaopotian.demo;

import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.FileInputStream;
import java.io.IOException;

/**
 * 文本提取
 * 如果要实现搜索、筛选功能，在一定场合下，比遍历整张表的单元格效率会高一些。
 */
public class test03 {
    public static void main(String[] args) throws IOException {
        FileInputStream fis = new FileInputStream("f:\\学生名单.xls");
        POIFSFileSystem fs = new POIFSFileSystem(fis);
        HSSFWorkbook wb = new HSSFWorkbook(fs);
        ExcelExtractor excelExtractor = new ExcelExtractor(wb);
        // 设置抽取的文本是否包括 Sheet 页的名称
        excelExtractor.setIncludeSheetNames(false);

        System.out.println(excelExtractor.getText());
        excelExtractor.close();
        wb.close();
        }
}
