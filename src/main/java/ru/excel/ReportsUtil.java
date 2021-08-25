package ru.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;

import java.util.*;

public class ReportsUtil {

    static Map<String, Row> createRowMap(String categoryName, Row row) {
        Map<String, Row> predictRowMap = new HashMap<>();
        predictRowMap.put(categoryName, row);
        return predictRowMap;
    }

    public static String getCategory(Row row) {
        Cell predictCell = row.getCell(ExcelReader.PREDICT_CELL_COLUMN_INDEX);
        String predictCellValue = "";
        switch (predictCell.getCellType()) {
            case STRING:
                //skip headers
                break;

            default:
                DataFormatter stringFormatter = new DataFormatter();
                predictCellValue = stringFormatter.formatCellValue(predictCell);
                break;
        }
        return predictCellValue;
    }

    public static String getCategory(Row row, int cellIndex) {
        Cell predictCell = row.getCell(cellIndex);
        String predictCellValue = "";
        switch (predictCell.getCellType()) {
            case STRING:
                //skip headers
                break;

            default:
                DataFormatter stringFormatter = new DataFormatter();
                predictCellValue = stringFormatter.formatCellValue(predictCell);
                break;
        }
        return predictCellValue;
    }

    public static boolean isMapHasTenRows(String predictNumOrName, List<Map<String, Row>> reportRowsList) {
        int rowCount = 0;
        for (Map<String, Row> filteredRow : reportRowsList) {
            for (Map.Entry<String, Row> predictRow : filteredRow.entrySet()) {
                String predict = predictRow.getKey();
                if (predictNumOrName.equals(predict)) {
                    rowCount++;
                }
            }
        }
        return rowCount == 10;
    }

    public static List<Map<String, Row>> createCategoryRowsList(List<Map<String, Row>> rowsList, String predict) {
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

    public static boolean isHeadersRow(String rowCategory) {
        return !rowCategory.equals("headers");
    }

    public static List<Map<String, Row>> createUniqueRowsList(List<Map<String, Row>> oldReportRowsList, List<Map<String, Row>> newReportRowsList, String category) {
        List<Map<String, Row>> uniqueRowsCategoryList = new ArrayList<>();
        for (Map<String, Row> newRowMap : newReportRowsList) {
            Row newRow = newRowMap.get(category);
            ArrayList<Boolean> rowContainsResult = new ArrayList<>();
            for (Map<String, Row> oldRowMap : oldReportRowsList) {
                Row oldRow = oldRowMap.get(category);
                if (newRow.equals(oldRow)) {
                    rowContainsResult.add(true);
                }
            }
            if (!rowContainsResult.contains(true)) {
                uniqueRowsCategoryList.add(newRowMap);
            }
        }
        return uniqueRowsCategoryList;
    }

    static List<Map<String, Row>> getUniqueTenRows(List<Map<String, Row>> uniqueRowsList) {
        List<Map<String, Row>> resultUniqueRowsList = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            resultUniqueRowsList.add(uniqueRowsList.get(i));

        }
        return resultUniqueRowsList;
    }

    static List<Map<String, Row>> appendUniqueRows(List<Map<String, Row>> uniqueRowsList, List<Map<String, Row>> categoryRowsList) {
        int firstUniqueElementIndex = getFirstUniqueElementIndex(uniqueRowsList, categoryRowsList);
        List<Map<String, Row>> resultRowsList = new ArrayList<>();

        for (int i = firstUniqueElementIndex; firstUniqueElementIndex > 0 && (uniqueRowsList.size() + resultRowsList.size() != 10); --i) {
            resultRowsList.add(0, categoryRowsList.get(i));
        }
        resultRowsList.addAll(uniqueRowsList);
        return sortResultRowList(resultRowsList);
    }

    private static List<Map<String, Row>> sortResultRowList(List<Map<String, Row>> resultRowsList) {
        List<Map<String, Row>> sortedMap = new ArrayList<>();
        Object[] resultRows = resultRowsList.toArray();
        bubbleSort(resultRows);
        for (Object sortedRow : resultRows) {
            sortedMap.add((Map<String, Row>) sortedRow);
        }
        return sortedMap;
    }

    private static void bubbleSort(Object[] resultRows) {
        for (int i = 0; i < resultRows.length - 1; i++) {
            for (int j = 0; j < resultRows.length - 1; j++) {
                Map<String, Row> firstMap = (Map<String, Row>) resultRows[j];
                int firstMapIdValue = getMapIdValue(firstMap);
                Map<String, Row> secondMap = (Map<String, Row>) resultRows[j + 1];
                int secondMapIdValue = getMapIdValue(secondMap);
                if (firstMapIdValue > secondMapIdValue) {
                    Object temp = resultRows[j];
                    resultRows[j] = resultRows[j + 1];
                    resultRows[j + 1] = temp;
                    break;
                }
            }
        }
    }

    private static int getMapIdValue(Map<String, Row> firstMap) {
        int idValue = -1;
        for (Map.Entry<String, Row> entry : firstMap.entrySet()) {
            Row firstRow = entry.getValue();
            idValue = Integer.parseInt(getCategory(firstRow, 0));
            break;
        }
        return idValue;
    }


    private static int getFirstUniqueElementIndex(List<Map<String, Row>> uniqueRowsList, List<Map<String, Row>> categoryRowsList) {
        int index = -1;
        for (int i = 0; i < categoryRowsList.size(); i++) {
            if (categoryRowsList.get(i).equals(uniqueRowsList.get(0))) {
                index = --i;
                break;
            }
        }
        return index;
    }
}
