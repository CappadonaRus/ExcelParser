package ru.excel;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static ru.excel.ExcelParser.parseExcelSheet;
import static ru.excel.ReportCategories.getCategoriesList;

public class FirstReport implements Reportable {

    private static List<Map<String, Row>> reportRowsList = new ArrayList<>();

    public static List<Map<String, Row>> getFirstReportRowsList() {
        return reportRowsList;
    }

    @Override
    public List<Map<String, Row>> createReport(XSSFWorkbook workBook, int sheetIndex) {
        List<String> categoriesList = getCategoriesList();
        XSSFSheet excelSheet = workBook.getSheetAt(sheetIndex);
        List<Map<String, Row>> excelRowsList = parseExcelSheet(excelSheet);
        return generateReport(categoriesList, excelRowsList);
    }


    public static List<Map<String, Row>> generateReport(List<String> categoriesList, List<Map<String, Row>> excelRowsList) {
        List<Map<String, Row>> reportList = new ArrayList<>();
        for (String category : categoriesList) {
            for (Map<String, Row> row : excelRowsList) {
                List<Map<String, Row>> categoryRowsList = createCategoryList(reportList, row, category);
                reportList.addAll(categoryRowsList);
            }
        }
        reportRowsList.addAll(reportList);
        return reportList;
    }

    private static List<Map<String, Row>> createCategoryList(List<Map<String, Row>> reportList, Map<String, Row> rowMap, String category) {
        List<Map<String, Row>> categoryRowsList = new ArrayList<>();
        for (Map.Entry<String, Row> entry : rowMap.entrySet()) {
            String rowCategory = entry.getKey();
            if (category.equals(rowCategory) && !ReportsUtil.isMapHasTenRows(category, reportList)) {
                Row excelRow = entry.getValue();
                Map<String, Row> excelRowMap = new HashMap<>();
                excelRowMap.put(rowCategory, excelRow);
                categoryRowsList.add(excelRowMap);
            }
        }
        return categoryRowsList;
    }
}
