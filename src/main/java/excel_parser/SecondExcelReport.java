package excel_parser;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.*;

import static excel_parser.ExcelWriter.*;

public class SecondExcelReport {

    public static void createSecondReport(XSSFWorkbook workBook) {
        List<Map<String, Row>> firstReportRowsList = FirstExcelReport.getFirstReportRowsList();
        XSSFSheet secondReportSheet = workBook.createSheet("Срез Ева без оператора");
        generateSecondReport(workBook, firstReportRowsList, secondReportSheet);
    }

    private static void generateSecondReport(XSSFWorkbook workBook, List<Map<String, Row>> firstReportRowsList, XSSFSheet createdSheet) {
        List<Map<String, Row>> secondReportRowsList = new ArrayList<>();
        List<Map<String, Row>> rowsWithoutClearAnswersList = ExcelReportsUtil.getRowsListWithoutAnswerCell(workBook.getSheetAt(DATA_SHEET_INDEX));
        List<String> predictList = getCategoriesList();

        for (String predict : predictList) {
            List<Map<String, Row>> categoryRowsList = ExcelReportsUtil.createCategoryRowsList(rowsWithoutClearAnswersList, predict);

            if (categoryRowsList.size() >= 20) {
                List<Map<String, Row>> uniqueRowsList = ExcelReportsUtil.createUniqueRowsList(firstReportRowsList, categoryRowsList);
                List<Map<String, Row>> shuffledTenRows = getShuffledTenRows(uniqueRowsList);
                secondReportRowsList.addAll(shuffledTenRows);
            } else if (categoryRowsList.size() >= 10) {
                List<Map<String, Row>> uniqueRowsList = ExcelReportsUtil.createUniqueRowsList(firstReportRowsList, categoryRowsList);
                if (uniqueRowsList.size() < 10) {
                    secondReportRowsList.addAll(uniqueRowsList);
                } else {
                    List<Map<String, Row>> shuffledTenRows = getShuffledTenRows(categoryRowsList);
                    secondReportRowsList.addAll(shuffledTenRows);
                }
            } else {
                secondReportRowsList.addAll(categoryRowsList);
            }
        }
        createReport(workBook, createdSheet, secondReportRowsList);
    }

    private static List<Map<String, Row>> getShuffledTenRows(List<Map<String, Row>> rowsWithoutAnswerList) {
        Collections.shuffle(rowsWithoutAnswerList);
        List<Map<String, Row>> categoryList = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            categoryList.add(rowsWithoutAnswerList.get(i));
        }
        return categoryList;
    }
}
