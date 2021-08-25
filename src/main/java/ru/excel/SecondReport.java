package ru.excel;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import static ru.excel.ExcelParser.parseExcelSheet;
import static ru.excel.ExcelParser.parseExcelSheetWithAnswerCell;
import static ru.excel.FirstReport.getFirstReportRowsList;
import static ru.excel.ReportCategories.getCategoriesList;
import static ru.excel.ReportsUtil.*;

public class SecondReport implements Reportable {

    private static List<Map<String, Row>> reportRowsList = new ArrayList<>();

    public static List<Map<String, Row>> getSecondReportRowsList() {
        return reportRowsList;
    }

    @Override
    public List<Map<String, Row>> createReport(XSSFWorkbook workBook, int sheetIndex) {
        List<Map<String, Row>> firstReportRowsList = getFirstReportRowsList();
        List<Map<String, Row>> excelRowsList = parseExcelSheetWithAnswerCell(workBook.getSheetAt(sheetIndex));
        return generateReport(firstReportRowsList, excelRowsList, getCategoriesList());
    }

    private static List<Map<String, Row>> generateReport(List<Map<String, Row>> firstReportList, List<Map<String, Row>> excelRowsList, List<String> categoriesList) {
        List<Map<String, Row>> secondReportList = new ArrayList<>();

        for (String category : categoriesList) {
            List<Map<String, Row>> firstReportCategoryList = createCategoryRowsList(firstReportList, category);
            List<Map<String, Row>> categoryList = createCategoryRowsList(excelRowsList, category);

            if (categoryList.size() >= 20) {
                List<Map<String, Row>> uniqueRowsList = createUniqueRowsList(firstReportCategoryList, categoryList, category);
                List<Map<String, Row>> uniqueTenRows = getUniqueTenRows(uniqueRowsList);
                secondReportList.addAll(uniqueTenRows);
            } else if (categoryList.size() >= 10) {
                List<Map<String, Row>> uniqueRowsList = createUniqueRowsList(firstReportCategoryList, categoryList, category);
                if (uniqueRowsList.size() >= 10) {
                    List<Map<String, Row>> uniqueTenRows = getUniqueTenRows(uniqueRowsList);
                    secondReportList.addAll(uniqueTenRows);
                } else {
                    List<Map<String, Row>> appendedUniqueRows = appendUniqueRows(uniqueRowsList, categoryList);
                    secondReportList.addAll(appendedUniqueRows);
                }
            } else {
                secondReportList.addAll(categoryList);
            }
        }
        return secondReportList;
    }
}