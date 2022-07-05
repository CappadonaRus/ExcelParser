package ru.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.*;

import static ru.excel.ExcelWriter.ANSWER_CELL_COLUMN_INDEX;
import static ru.excel.ReportsUtil.createRowMap;

public class ExcelParser {

    public static List<Map<String, Row>> parseExcelSheet(Sheet sheet) {
        return parseExcelSheet(sheet, ReportType.FIRST);
    }

    public static List<Map<String, Row>> parseExcelSheetWithAnswerCell(Sheet sheet) {
        return parseExcelSheet(sheet, ReportType.SECOND);
    }

    private static List<Map<String, Row>> parseExcelSheet(Sheet sheet, ReportType reportType) {
        List<Map<String, Row>> resultRowsList = new ArrayList<>();
        Iterator<Row> rowIterator = sheet.iterator();
        Row headersRow = rowIterator.next();
        Map<String, Row> headersMap = createRowMap("headers", headersRow);
        resultRowsList.add(headersMap);
        while (rowIterator.hasNext()) {
            Row currentRow = rowIterator.next();
            Map<String, Row> rowMap  = getStringRowMap(reportType, currentRow);
            resultRowsList.add(rowMap);
        }
        return resultRowsList;
    }

    private static Map<String, Row> getStringRowMap(ReportType reportType, Row currentRow) {
        Map<String, Row> rowMap = new HashMap<>();
        switch (reportType) {
            case FIRST:
                String category = ReportsUtil.getCategory(currentRow);
                rowMap = createRowMap(category, currentRow);
                break;
            case SECOND:
                Cell answerCell = currentRow.getCell(ANSWER_CELL_COLUMN_INDEX);
                if (answerCell != null) {
                    String predictCategoryNum = ReportsUtil.getCategory(currentRow);
                    rowMap = createRowMap(predictCategoryNum, currentRow);
                }
                break;
        }
        return rowMap;
    }
}
