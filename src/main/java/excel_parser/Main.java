package excel_parser;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import static excel_parser.ExcelReportsUtil.createFirstReport;
import static excel_parser.FirstExcelReport.createReportCategoriesList;
import static excel_parser.SecondExcelReport.createSecondReport;

public class Main {

    private static String filePath = "./discharge (36).xlsx";

    public static String getFilePath() {
        return filePath;
    }


    public static void main(String[] args) {
        XSSFWorkbook workBook = ExcelReader.readExcelWorkBook(filePath);
        Sheet dataSheet = workBook.getSheetAt(0);
        ExcelReportsUtil.getRowsListWithoutAnswerCell(dataSheet);
        createReportCategoriesList();
        createFirstReport(workBook);
        createSecondReport(workBook);
        ExcelWriter.writeSheetIntoBook(workBook);
    }
}
