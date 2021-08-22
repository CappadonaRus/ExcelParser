package ru.excel;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Map;

import static ru.excel.ExcelReportsUtil.getPredictCellValue;
import static ru.excel.ExcelWriter.*;

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
                List<Map<String, Row>> uniqueRowsList = ExcelReportsUtil.createUniqueRowsList(oldCategoryRowsList, categoryRowsList, category);
                List<Map<String, Row>> uniqueTenRows = getUniqueTenRows(uniqueRowsList);
                secondReportRowsList.addAll(uniqueTenRows);
            } else if (categoryRowsList.size() >= 10) {
                List<Map<String, Row>> uniqueRowsList = ExcelReportsUtil.createUniqueRowsList(oldCategoryRowsList, categoryRowsList, category);
                if (uniqueRowsList.size() >= 10) {
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
            idValue = Integer.parseInt(getPredictCellValue(firstRow, 0));
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

    private static List<Map<String, Row>> getShuffledTenRows(List<Map<String, Row>> rowsWithoutAnswerList) {
        Collections.shuffle(rowsWithoutAnswerList);
        List<Map<String, Row>> categoryList = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            categoryList.add(rowsWithoutAnswerList.get(i));
        }
        return categoryList;
    }
}