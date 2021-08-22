package ru.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.*;

import static ru.excel.ExcelWriter.ANSWER_CELL_COLUMN_INDEX;
import static ru.excel.ExcelWriter.createReport;

public class ExcelReportsUtil {

    private static List<Map<String, Row>> filteredRowsForReport = new ArrayList<>();

    public static List<Map<String, Row>> getFilteredRowsForReport() {
        return filteredRowsForReport;
    }

    public static void createFirstReport(XSSFWorkbook workBook) {
        List<Map<String, Row>> filteredRowsList = FirstExcelReport.getFirstReportRowsList();
        XSSFSheet sheet = workBook.createSheet("���� ���");
        createReport(workBook, sheet, filteredRowsList);
    }

    public static List<Map<String, Row>> getRowsListWithoutAnswerCell(Sheet dataSheet) {
        List<Map<String, Row>> rowsListWithoutAnswer = new ArrayList<>();
        Iterator<Row> rowIterator = dataSheet.iterator();
        Row headersRow = rowIterator.next();
        Map<String, Row> headersMap = createNamedRowMap("headers", headersRow);
        rowsListWithoutAnswer.add(headersMap);

        while (rowIterator.hasNext()) {
            Row currentRow = rowIterator.next();
            Cell answerCell = currentRow.getCell(ANSWER_CELL_COLUMN_INDEX);
            if (answerCell != null) {
                String predictCategoryNum = ExcelReportsUtil.getPredictCellValue(currentRow);
                Map<String, Row> predictMap = createNamedRowMap(predictCategoryNum, currentRow);
                rowsListWithoutAnswer.add(predictMap);
            }
        }
        filteredRowsForReport.addAll(rowsListWithoutAnswer);
        return rowsListWithoutAnswer;
    }

    public static List<Map<String, Row>> getRowsListWithOperatorAnswerCell(Sheet dataSheet) {
        List<Map<String, Row>> rowsWithOperatorAnswerList = new ArrayList<>();
        Iterator<Row> rowIterator = dataSheet.iterator();
        Row headersRow = rowIterator.next();
        Map<String, Row> headersMap = createNamedRowMap("headers", headersRow);
        rowsWithOperatorAnswerList.add(headersMap);

        while (rowIterator.hasNext()) {
            Row currentRow = rowIterator.next();
            String predictCategoryNum = ExcelReportsUtil.getPredictCellValue(currentRow);
            Map<String, Row> predictMap = createNamedRowMap(predictCategoryNum, currentRow);
            rowsWithOperatorAnswerList.add(predictMap);
        }
        filteredRowsForReport.addAll(rowsWithOperatorAnswerList);
        return rowsWithOperatorAnswerList;
    }

    private static Map<String, Row> createNamedRowMap(String predictNumOrName, Row row) {
        Map<String, Row> predictRowMap = new HashMap<>();
        predictRowMap.put(predictNumOrName, row);
        return predictRowMap;
    }

    public static String getPredictCellValue(Row row) {
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

    public static String getPredictCellValue(Row row, int cellIndex) {
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
}