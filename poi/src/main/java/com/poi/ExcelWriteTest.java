package com.poi;


import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;

public class ExcelWriteTest {
    String PATH = "C:\\Users\\Atlantis\\Documents\\Ideaspace\\poi-study\\";
    @Test
    public void testXLS() throws Exception {
        //1.创建一个工作簿（这个是XLS格式，最多支持65536条，但是速度快，因为过程中写入缓存，不操作磁盘，最后一次性写入到磁盘）
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
        //1.创建一个工作簿（这个是XLSX格式，这是大文件数据，但是速度慢，因为这个是写入到内存中去，但是可以写入大数据）
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
    @Test
    public void testBigData03() throws Exception {
        long begin = System.currentTimeMillis();

        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet("这是我创建用来测试填充速度的表");
        for (int i = 0; i < 65536; i++) {
            Row row = sheet.createRow(i);
            for (int j = 0; j < 10; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(j);
            }
        }
        FileOutputStream outputStream = new FileOutputStream(PATH + "测试文件速度的HSSF格式(65536).xls");
        workbook.write(outputStream);
        outputStream.close();
        long end = System.currentTimeMillis();
        //这里除以1000是因为记录的是毫秒值
        double time = ((double)(end-begin)/1000);
        System.out.println(time + "秒");
    }
    @Test
    public void testBigData07() throws Exception {
        long begin = System.currentTimeMillis();
        //使用这个格式的速度是比较慢的，但是可以存储大数据
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("这是我创建用来测试填充速度的表");
        for (int i = 0; i < 100000; i++) {
            Row row = sheet.createRow(i);
            for (int j = 0; j < 10; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(j);
            }
        }
        FileOutputStream outputStream = new FileOutputStream(PATH + "测试文件速度的XSSF格式(65536).xlsx");
        workbook.write(outputStream);
        outputStream.close();
        long end = System.currentTimeMillis();
        //这里除以1000是因为记录的是毫秒值
        double time = ((double)(end-begin)/1000);
        System.out.println(time + "秒");
    }
    @Test
    public void testBigData07SXSSF() throws Exception {
        long begin = System.currentTimeMillis();
        //SXSSF格式是优化版的，速度更快
        //原理就是在插入的时候会有一个缓存文件，当插入的数据达到100条的时候，会先将数据写入缓存
        Workbook workbook = new SXSSFWorkbook();
        Sheet sheet = workbook.createSheet("这是我创建用来测试填充速度的表");
        for (int i = 0; i < 100000; i++) {
            Row row = sheet.createRow(i);
            for (int j = 0; j < 10; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(j);
            }
        }
        FileOutputStream outputStream = new FileOutputStream(PATH + "测试文件速度的XSSF格式(65536).xlsx");
        workbook.write(outputStream);
        outputStream.close();
        //这里有一个坑，强转的时候这里有一个空格
        ((SXSSFWorkbook) workbook).dispose();
        long end = System.currentTimeMillis();
        //这里除以1000是因为记录的是毫秒值
        double time = ((double)(end-begin)/1000);
        System.out.println(time + "秒");
    }

}
