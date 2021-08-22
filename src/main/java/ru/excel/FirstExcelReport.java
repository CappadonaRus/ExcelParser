package ru.excel;

import org.apache.poi.ss.usermodel.Row;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class FirstExcelReport {

    private static List<Map<String, Row>> firstReportRowsList = new ArrayList<>();

    public static List<Map<String, Row>> getFirstReportRowsList() {
        return firstReportRowsList;
    }

    public static void createReportCategoriesList() {
        List<Map<String, Row>> filteredRowsList = ExcelReportsUtil.getFilteredRowsForReport();
        for (String predictsList : ExcelWriter.getCategoriesList()) {
            for (Map<String, Row> predictRow : filteredRowsList) {
                List<Map<String, Row>> categoryRowsList = createCategoryRowsList(predictRow, predictsList);
                firstReportRowsList.addAll(categoryRowsList);
            }
        }
    }

    private static List<Map<String, Row>> createCategoryRowsList(Map<String, Row> predictRow, String predictFilterValue) {
        List<Map<String, Row>> filteredMapsList = new ArrayList<>();
        for (Map.Entry<String, Row> row : predictRow.entrySet()) {
            String predictValue = row.getKey();
            if (predictFilterValue.equals(predictValue) && !ExcelReportsUtil.isMapHasTenRows(predictFilterValue, firstReportRowsList)) {
                Row excelRow = row.getValue();
                Map<String, Row> foundPredictMap = new HashMap<>();
                foundPredictMap.put(predictValue, excelRow);
                filteredMapsList.add(foundPredictMap);
            }
        }
        return filteredMapsList;
    }

}
