package ru.excel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;


public class ExcelReader {

    public static final int PREDICT_CELL_COLUMN_INDEX = 5;
    private static XSSFWorkbook workbook;


    public static XSSFWorkbook getWorkBook() {
        return workbook;
    }

    public static XSSFWorkbook readExcelWorkBook(String filePath) {
        try (FileInputStream inputStream = new FileInputStream(new File(filePath))) {
            workbook = new XSSFWorkbook(inputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return workbook;
    }
}
