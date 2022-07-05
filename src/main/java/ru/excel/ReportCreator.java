package ru.excel;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.List;
import java.util.Map;

import static ru.excel.ExcelWriter.writeReportIntoSheet;
import static ru.excel.ExcelWriter.writeSheetIntoBook;

public class ReportCreator {
    public static final int FIRST_SHEET_WORKBOOK_INDEX = 0;

    private static String filePath = "./eva_new.xlsx";

//    private static String filePath = "";

    public static String getFilePath() {
        return filePath;
    }


    public static void main(String[] args) {
//        filePath = args[0] + ".xlsx";
        XSSFWorkbook workBook = ExcelReader.readExcelWorkBook(filePath);

        FirstReport firstReport = new FirstReport();
        SecondReport secondReport = new SecondReport();

        List<Map<String, Row>> firstReportResultList = firstReport.createReport(workBook, FIRST_SHEET_WORKBOOK_INDEX);
        List<Map<String, Row>> secondReportResultList = secondReport.createReport(workBook, FIRST_SHEET_WORKBOOK_INDEX);
        fillSheetAndWrite(workBook, firstReportResultList, "Срез Ева");
        fillSheetAndWrite(workBook, secondReportResultList, "Срез Ева без оператора");
    }

    private static void fillSheetAndWrite(XSSFWorkbook workBook, List<Map<String, Row>> reportList, String sheetName) {
        XSSFSheet firstSheet = workBook.createSheet(sheetName);
        writeReportIntoSheet(workBook, firstSheet, reportList);
        writeSheetIntoBook(workBook);
    }
}
