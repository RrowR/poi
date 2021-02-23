package com.poi;


import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileOutputStream;

public class ExcelWriteTest {
    String PATH = "C:\\Users\\Atlantis\\Documents\\Ideaspace\\poi-study\\poi";
    @Test
    public void testXLS() throws Exception {
        //1.创建一个工作簿（这个是XLS格式）
        Workbook workbook = new HSSFWorkbook();
        //2.创建一个工作表（这个是有返回值的）
        Sheet sheet = workbook.createSheet("这是代码创建的一张表");
        //3.创建一行
        Row row = sheet.createRow(0);
        //4.创建一列
        Cell cell = row.createCell(0);
        //将数据添加到这一个单元格中
        cell.setCellValue("这是第一列第一行的内容");

        //将数据添加到第二行第一列的单元格中
        Row row1 = sheet.createRow(1);
        Cell cell1 = row1.createCell(0);
        cell1.setCellValue("这是第二行内容");

        FileOutputStream outputStream = new FileOutputStream(PATH + "第一个文件(只能建立65536个行的文件格式).xls");
        workbook.write(outputStream);
        outputStream.close();
    }

    @Test
    public void testXLSX() throws Exception {
        //1.创建一个工作簿（这个是XLSX格式）
        Workbook workbook = new XSSFWorkbook();
        //2.创建一个工作表（这个是有返回值的）
        Sheet sheet = workbook.createSheet("这是代码创建的一张表");
        //3.创建一行
        Row row = sheet.createRow(0);
        //4.创建一列
        Cell cell = row.createCell(0);
        //将数据添加到这一个单元格中
        cell.setCellValue("111");

        //将数据添加到第二行第一列的单元格中
        Row row1 = sheet.createRow(1);
        Cell cell1 = row1.createCell(0);
        String time = new DateTime().toString("yyyy-MM-dd HH:mm:SS");
        cell1.setCellValue(time);

        FileOutputStream outputStream = new FileOutputStream(PATH + "第2个文件(只能建立65536个行的文件格式).xlsx");
        workbook.write(outputStream);
        outputStream.close();
    }
}
