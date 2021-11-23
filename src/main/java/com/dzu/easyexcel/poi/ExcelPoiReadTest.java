package com.dzu.easyexcel.poi;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Controller;

import javax.annotation.PostConstruct;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

/**
 * @author by ZhaoDong
 * @Classname ExcelPoiReadTest
 * @Description 使用 poi 读 excel
 * @Date 2021/11/23 22:41
 */
@Controller
public class ExcelPoiReadTest {
    private static final String PATH = "E:\\JAVA\\easy-excel\\excel\\";

    @PostConstruct
    public void testPoiReadAll() {
        XSSFWorkbook workbook = getWorkbook();
        if (workbook == null) {
            return;
        }
        // 获取sheet的个数
        int sheets = workbook.getNumberOfSheets();
        // 获取第一个sheet
        XSSFSheet sheet = workbook.getSheetAt(0);
        // 获取所有的行
        int rowCount = sheet.getPhysicalNumberOfRows();
        // 获取第一行
        XSSFRow row = sheet.getRow(0);

        // 第一行所有的列
        int cellCount = row.getPhysicalNumberOfCells();
        for (int i = 0; i < cellCount; i++) {
            XSSFCell cell = row.getCell(i);
            // 获取数据类型
            CellType cellTypeEnum = cell.getCellTypeEnum();
            switch (cellTypeEnum) {
                case STRING:
                    System.out.println(cellTypeEnum);
                    break;
                case NUMERIC:
                    System.out.println(cellTypeEnum);
                    break;
                default:
                    System.out.println("213");
            }
            // 获取数据
            String stringCellValue = cell.getStringCellValue();
        }

    }

    public void testPoiRead() {
        XSSFWorkbook workbook = getWorkbook();
        if (workbook == null) {
            return;
        }
        XSSFSheet sheet = workbook.getSheetAt(0);
        System.out.println(workbook.getNumberOfSheets());
        XSSFRow row = sheet.getRow(0);
        XSSFCell cell = row.getCell(0);
        String stringCellValue = cell.getStringCellValue();
        System.out.println(stringCellValue);
    }

    private XSSFWorkbook getWorkbook() {
        FileInputStream inputStreamReader = null;
        try {
            inputStreamReader = new FileInputStream(PATH + "03.xls");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        XSSFWorkbook workbook = null;
        try {
            if (inputStreamReader == null) {
                return null;
            }
            workbook = new XSSFWorkbook(inputStreamReader);
        } catch (IOException e) {
            e.printStackTrace();
        }
        if (workbook == null) {
            return null;
        }
        return workbook;
    }
}
