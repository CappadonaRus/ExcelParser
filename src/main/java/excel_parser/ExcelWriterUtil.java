package excel_parser;


import org.apache.poi.ss.usermodel.Row;

import java.util.*;

public class ExcelWriterUtil {


    private static List<String> predictList = new ArrayList<>();
    private static List<Map<String, Row>> rowsForFirstReport = new ArrayList<>();

    static {
        predictList.add("headers");
        for (int i = 1; i <= 125; i++) {
            predictList.add(String.valueOf(i));
        }
    }

    public static List<Map<String, Row>> getFilteredMap() {
        return rowsForFirstReport;
    }

    public static List<String> getPredictList() {
        return predictList;
    }


    public static void filterRows() {
        List<Map<String, Row>> excelRowsMap = ExcelReader.getExcelRowMap();
        Collections.shuffle(excelRowsMap);
        for (String predictsList : predictList) {
            for (Map<String, Row> predictRow : excelRowsMap) {
                createMapWithFilteredRows(predictRow, predictsList);
            }
        }
    }

    private static void createMapWithFilteredRows(Map<String, Row> predictRow, String predictFilterValue) {
        for (Map.Entry<String, Row> row : predictRow.entrySet()) {
            String predictValue = row.getKey();
            if (predictFilterValue.equals(predictValue) && !isMapHasTenRows(predictFilterValue)) {
                Row excelRow = row.getValue();
                Map<String, Row> foundPredictMap = new HashMap<>();
                foundPredictMap.put(predictValue, excelRow);
                rowsForFirstReport.add(foundPredictMap);
            }
        }
    }


    private static boolean isMapHasTenRows(String predictFilter) {
        int rowCount = 0;
        for (Map<String, Row> filteredRow : rowsForFirstReport) {
            for (Map.Entry<String, Row> predictRow : filteredRow.entrySet()) {
                String predict = predictRow.getKey();
                if (predictFilter.equals(predict)) {
                    rowCount++;
                }
            }
        }
        return rowCount == 10;
    }

    static List<Map<String, Row>> filterListForPredict(List<Map<String, Row>> rowsList, String predict) {
        List<Map<String, Row>> filteredList = new ArrayList<>();
        for (Map<String, Row> rowMap : rowsList) {
            for (Map.Entry<String, Row> row : rowMap.entrySet()) {
                String predictNumber = row.getKey();
                if (predictNumber.equals(predict)) {
                    filteredList.add(rowMap);
                }
            }
        }
        return filteredList;
    }
}
