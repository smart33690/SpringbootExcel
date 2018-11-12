package com.xiaopotian.demo;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;

import java.io.*;

/**
 * 读取和重写工作簿
 * 需求：我们读取一张工作簿里面的一张表，然后修改其中的单元格数据。
 * 思路：先把一份文件读取进内存，然后先 get 再 set ，这样的操作就完成了修改，然后再全部写回。
 */
public class test09 {
    public static void main(String[] args) throws IOException {
        InputStream is = new FileInputStream("f:\\学生名单.xls");
        POIFSFileSystem fs = new POIFSFileSystem(is);
        Workbook wb = new HSSFWorkbook(fs);

        // 获取第 1 个 Sheet 页
        Sheet sheet = wb.getSheetAt(0);
        // 获取第 1 行
        Row row = sheet.getRow(0);
        // 获取 Cell单元格
        Cell cell = row.getCell(5);
        if (cell == null) {
            cell = row.createCell(3);
            cell.setCellValue("没有读取到数据");
        }else {
            cell.setCellType(Cell.CELL_TYPE_STRING);
            cell.setCellValue("大侠高姓大名啊啊啊嗯分嗯嗯");
        }

        FileOutputStream fos;

        try {
            fos = new FileOutputStream("f:\\学生名单.xls");
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
