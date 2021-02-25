package com.poi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.IOException;


public class ExcelReadTest {
    String PATH = "C:\\Users\\Atlantis\\Documents\\Ideaspace\\poi-study\\";
    @Test
    public void testExcelRead07() throws IOException {
        //如果想得到表的话，这里需要使用到输入流，而输入流需要传入到下面的表对象里
        FileInputStream fileInputStream = new FileInputStream(PATH + "第2个文件(只能建立65536个行的文件格式).xlsx");
        Workbook workbook = new XSSFWorkbook(fileInputStream);
        Sheet sheet = workbook.getSheetAt(0);
        Row row = sheet.getRow(1);
        Cell cell = row.getCell(0);
        fileInputStream.close();
        System.out.println(cell.getStringCellValue());
    }

    @Test
    public void testExcelRead03() throws IOException {
        //如果想得到表的话，这里需要使用到输入流，而输入流需要传入到下面的表对象里
        FileInputStream fileInputStream = new FileInputStream(PATH + "测试文件速度的HSSF格式(65536).xls");
        Workbook workbook = new HSSFWorkbook(fileInputStream);
        Sheet sheet = workbook.getSheetAt(0);
        Row row = sheet.getRow(1);
        Cell cell = row.getCell(0);
        fileInputStream.close();
        //getStringCellValue获取字符串，这个现象只出现在HSSF中也就是xls这个类型里
        System.out.println(cell.getNumericCellValue());
    }

}
