package ru.excel;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReportCreator {

    private static String filePath = "./eva.xlsx";

//    private static String filePath = "";

    public static String getFilePath() {
        return filePath;
    }


    public static void main(String[] args) {
//        filePath = args[0] + ".xlsx";
        XSSFWorkbook workBook = ExcelReader.readExcelWorkBook(filePath);
        Sheet dataSheet = workBook.getSheetAt(0);
        ExcelReportsUtil.getRowsListWithOperatorAnswerCell(dataSheet);
        FirstExcelReport.createReportCategoriesList();
        ExcelReportsUtil.createFirstReport(workBook);
        SecondExcelReport.createSecondReport(workBook);
        ExcelWriter.writeSheetIntoBook(workBook);
    }
}
