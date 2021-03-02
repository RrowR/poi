package com.poi;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Date;


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

    @Test
    public void testCellType() throws IOException {
        //使用输入流获取这个文件里的属性
        FileInputStream inputStream = new FileInputStream(PATH + "明细表.xlsx");
        //使用XSSFWorkbook对象来获得xlsx格式文件
        Workbook workbook = new XSSFWorkbook(inputStream);
        //获取当前文件里的一个sheet，就是左下角的第几个表
        Sheet sheet = workbook.getSheetAt(0);
        //获取第一行
        Row row = sheet.getRow(0);
        if (row != null){       //如果当前的行不为空
            //获取每一行的个数
            int cellsCount = row.getPhysicalNumberOfCells();
            //进行循环判断，让这里的每一条数据都被遍历出来
            for (int cellNum = 0; cellNum < cellsCount ; cellNum++) {
                //获取每一列的数据
                Cell cell = row.getCell(cellNum);
                //如果当前列的数据不等于空
                if (cell != null){
                    //将这一列的数据转换为字符串类型，并进行输出打印
                    String cellValue = cell.getStringCellValue();
                    System.out.print(cellValue + " | ");
                }
            }
            System.out.println();
        }

        //获取每一 行 的数据个数
        int rowCount = sheet.getPhysicalNumberOfRows();
        for (int rowNum = 1; rowNum < rowCount; rowNum++) {
            //获取遍历出来的当前行的数据
            Row rowData = sheet.getRow(rowNum);
            //获取当前行的数据如果不为空
            if (rowData != null){
                //读取当前行每一列的数据
                int cells = rowData.getPhysicalNumberOfCells();
                for (int cellNum = 0; cellNum < cells; cellNum++) {
                    System.out.print("[" + (rowNum+1) + "-" + (cellNum+1) + "]");
                    //获取每一列的数据,这里还是从第0列开始获取
                    Cell cell = rowData.getCell(cellNum);
                    //匹配列的数据类型
                    if (cell != null){
                        int cellType = cell.getCellType();      //获取当前列的类型
                        String cellValue = "";
                        switch (cellType){
                            case HSSFCell.CELL_TYPE_STRING:     //字符串类型
                                System.out.print("[String]");
                                cellValue = cell.getStringCellValue();
                                break;
                            case HSSFCell.CELL_TYPE_BOOLEAN:     //布尔类型
                                System.out.print("[Boolean]");
                                cellValue = String.valueOf(cell.getBooleanCellValue());
                                break;
                            case HSSFCell.CELL_TYPE_BLANK:     //空
                                System.out.print("[null]");
                                break;
                            case HSSFCell.CELL_TYPE_NUMERIC:     //数字类型，还要继续判断是什么数字类型
                                System.out.print("[NUMERIC]");
                                if (HSSFDateUtil.isCellDateFormatted(cell)){
                                    System.out.print("【日期】");
                                    Date date = cell.getDateCellValue();
                                    cellValue = new DateTime(date).toString("yyyy-MM-dd");
                                }else {
                                    //如果不是日期格式，防止数字过长
                                    System.out.print("【转换为字符串进行输出】");
                                    cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                                    cellValue = cell.toString();
                                }
                                break;
                            case HSSFCell.CELL_TYPE_ERROR:     //空
                                System.out.print("[类型错误]");
                                break;
                        }
                        System.out.println(cellValue);
                    }
                }
            }
        }
        inputStream.close();
    }

    //抽取方法
    public void fileTranform(FileInputStream inputStream) throws IOException {
        //使用XSSFWorkbook对象来获得xlsx格式文件
        Workbook workbook = new XSSFWorkbook(inputStream);
        //获取当前文件里的一个sheet，就是左下角的第几个表
        Sheet sheet = workbook.getSheetAt(0);
        //获取第一行
        Row row = sheet.getRow(0);
        if (row != null){       //如果当前的行不为空
            //获取每一行的个数
            int cellsCount = row.getPhysicalNumberOfCells();
            //进行循环判断，让这里的每一条数据都被遍历出来
            for (int cellNum = 0; cellNum < cellsCount ; cellNum++) {
                //获取每一列的数据
                Cell cell = row.getCell(cellNum);
                //如果当前列的数据不等于空
                if (cell != null){
                    //将这一列的数据转换为字符串类型，并进行输出打印
                    String cellValue = cell.getStringCellValue();
                    System.out.print(cellValue + " | ");
                }
            }
            System.out.println();
        }

        //获取每一 行 的数据个数
        int rowCount = sheet.getPhysicalNumberOfRows();
        for (int rowNum = 1; rowNum < rowCount; rowNum++) {
            //获取遍历出来的当前行的数据
            Row rowData = sheet.getRow(rowNum);
            //获取当前行的数据如果不为空
            if (rowData != null){
                //读取当前行每一列的数据
                int cells = rowData.getPhysicalNumberOfCells();
                for (int cellNum = 0; cellNum < cells; cellNum++) {
                    System.out.print("[" + (rowNum+1) + "-" + (cellNum+1) + "]");
                    //获取每一列的数据,这里还是从第0列开始获取
                    Cell cell = rowData.getCell(cellNum);
                    //匹配列的数据类型
                    if (cell != null){
                        int cellType = cell.getCellType();      //获取当前列的类型
                        String cellValue = "";
                        switch (cellType){
                            case HSSFCell.CELL_TYPE_STRING:     //字符串类型
                                System.out.print("[String]");
                                cellValue = cell.getStringCellValue();
                                break;
                            case HSSFCell.CELL_TYPE_BOOLEAN:     //布尔类型
                                System.out.print("[Boolean]");
                                cellValue = String.valueOf(cell.getBooleanCellValue());
                                break;
                            case HSSFCell.CELL_TYPE_BLANK:     //空
                                System.out.print("[null]");
                                break;
                            case HSSFCell.CELL_TYPE_NUMERIC:     //数字类型，还要继续判断是什么数字类型
                                System.out.print("[NUMERIC]");
                                if (HSSFDateUtil.isCellDateFormatted(cell)){
                                    System.out.print("【日期】");
                                    Date date = cell.getDateCellValue();
                                    cellValue = new DateTime(date).toString("yyyy-MM-dd");
                                }else {
                                    //如果不是日期格式，防止数字过长
                                    System.out.print("【转换为字符串进行输出】");
                                    cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                                    cellValue = cell.toString();
                                }
                                break;
                            case HSSFCell.CELL_TYPE_ERROR:     //空
                                System.out.print("[类型错误]");
                                break;
                        }
                        System.out.println(cellValue);
                    }
                }
            }
        }
        inputStream.close();
    }
}

