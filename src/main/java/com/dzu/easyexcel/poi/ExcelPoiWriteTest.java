package com.dzu.easyexcel.poi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.springframework.stereotype.Controller;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @author by ZhaoDong
 * @Classname ExcelPoiTest
 * @Description poi 操作excel
 * HSSF提供读写Microsoft Excel XLS(03)格式档案的功能。
 * <p>
 * XSSF提供读写Microsoft Excel OOXML XLSX(07)格式档案的功能。
 * <p>
 * HWPF提供读写Microsoft Word DOC格式档案的功能。
 * <p>
 * HSLF提供读写Microsoft PowerPoint格式档案的功能。
 * <p>
 * HDGF提供读Microsoft Visio格式档案的功能。
 * <p>
 * HPBF提供读Microsoft Publisher格式档案的功能。
 * <p>
 * HSMF提供读Microsoft Outlook格式档案的功能。
 * @Date 2021/11/22 21:48
 */
@Controller
public class ExcelPoiWriteTest {

    private static final String PATH = "E:\\JAVA\\easy-excel\\excel\\";


    /**
     * 测试 优化版
     */
    public void testPoiSXSSFWorkbook() {
        SXSSFWorkbook workbook = new SXSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        for (int i = 0; i < 100000; i++) {
            Row row = sheet.createRow(i);
            row.createCell(0).setCellValue(i);
        }
        FileOutputStream fileOutputStream = null;
        try {
            fileOutputStream = new FileOutputStream(PATH + "测试03.xlsx");
            // 输出
            workbook.write(fileOutputStream);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            // 关闭流
            try {
                if (fileOutputStream != null) {
                    fileOutputStream.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        workbook.dispose();
    }

    /**
     * OOM报错
     */
    public void testPoi03OOM() {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet();
        for (int i = 0; i < Integer.MAX_VALUE; i++) {
            XSSFRow row = sheet.createRow(i);
            row.createCell(0).setCellValue(i);
        }
        FileOutputStream fileOutputStream = null;
        try {
            fileOutputStream = new FileOutputStream(PATH + "测试03.xlsx");
            // 输出
            workbook.write(fileOutputStream);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            // 关闭流
            try {
                if (fileOutputStream != null) {
                    fileOutputStream.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

    }

    /**
     * 03版和07版excel 基本操作
     */
    public void testPoi03() {
        // 1、创建工作薄

        // 07版本使用
        Workbook workbook07 = new XSSFWorkbook();
        // 03版本使用
        Workbook workbook = new HSSFWorkbook();
        // 2、创建工作表
        Sheet sheet = workbook.createSheet("测试用表");
        // 3、创建行
        Row row = sheet.createRow(0);
        // 4、创建单元格 0,0
        Cell cell = row.createCell(0);
        cell.setCellValue("今天学了多久？");

        // 5、 0，1
        Cell cell1 = row.createCell(1);
        cell1.setCellValue(66);

        // 6、第二行
        Row row1 = sheet.createRow(1);
        Cell cell2 = row1.createCell(0);
        cell2.setCellValue("日期：");
        String time = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        row1.createCell(1).setCellValue(time);

        // 7、生成表
        // IO流，03版本
        FileOutputStream fileOutputStream = null;
        try {
            fileOutputStream = new FileOutputStream(PATH + "测试03.xls");
            // 输出
            workbook.write(fileOutputStream);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            // 关闭流
            try {
                if (fileOutputStream != null) {
                    fileOutputStream.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        System.out.println("okk");

    }

}
