package excel_parser;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Map;

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
        List<String> categoryList = getCategoriesList();

        for (String category : categoryList) {
            List<Map<String, Row>> categoryRowsList = ExcelReportsUtil.createCategoryRowsList(rowsWithoutClearAnswersList, category);
            List<Map<String, Row>> oldCategoryRowsList = ExcelReportsUtil.createCategoryRowsList(firstReportRowsList, category);

            if (categoryRowsList.size() >= 20) {
                List<Map<String, Row>> uniqueRowsList = ExcelReportsUtil.createUniqueRowsList(oldCategoryRowsList, categoryRowsList,category);
                List<Map<String, Row>> uniqueTenRows = getUniqueTenRows(uniqueRowsList);
//                List<Map<String, Row>> shuffledTenRows = getShuffledTenRows(uniqueRowsList);
                secondReportRowsList.addAll(uniqueTenRows);
            } else if (categoryRowsList.size() >= 10) {
                List<Map<String, Row>> uniqueRowsList = ExcelReportsUtil.createUniqueRowsList(oldCategoryRowsList, categoryRowsList,category);
                if (uniqueRowsList.size() > 10) {
                    List<Map<String, Row>> uniqueTenRows = getUniqueTenRows(uniqueRowsList);
                    secondReportRowsList.addAll(uniqueTenRows);
                } else {
                    List<Map<String, Row>> appendedUniqueRows = appendUniqueRows(uniqueRowsList, categoryRowsList);
                    secondReportRowsList.addAll(appendedUniqueRows);
                }
            } else {
                secondReportRowsList.addAll(categoryRowsList);
            }
        }
        createReport(workBook, createdSheet, secondReportRowsList);
    }

    private static List<Map<String, Row>> getUniqueTenRows(List<Map<String, Row>> uniqueRowsList) {
        List<Map<String, Row>> resultUniqueRowsList = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            resultUniqueRowsList.add(uniqueRowsList.get(i));

        }
        return resultUniqueRowsList;
    }

    private static List<Map<String, Row>> appendUniqueRows(List<Map<String, Row>> uniqueRowsList, List<Map<String, Row>> categoryRowsList) {
        List<Map<String, Row>> resultRowsList = new ArrayList<>(uniqueRowsList);
        int uniqueRowsSize = uniqueRowsList.size();
        for (int i = uniqueRowsSize; i < 10; i++) {
            resultRowsList.add(categoryRowsList.get(i));
        }
        return resultRowsList;
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