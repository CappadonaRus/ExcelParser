package excel_parser;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.*;

import static excel_parser.ExcelReader.getPredictCellValue;
import static excel_parser.ExcelReader.getWorkBook;
import static excel_parser.ExcelWriter.*;
import static excel_parser.ExcelWriterUtil.*;

public class SecondExcelReport {

    public static void createSecondReport(XSSFWorkbook workBook) {
        List<Map<String, Row>> firstReportRowsList = getFilteredMap();
        XSSFSheet sheet = workBook.createSheet("Срез Ева без оператора");
        generateSecondReport(workBook, firstReportRowsList, sheet);
    }

    private static void generateSecondReport(XSSFWorkbook workBook, List<Map<String, Row>> firstReportRowsList, XSSFSheet sheet) {
        List<Map<String, Row>> rowsWithoutClearAnswersList = filterClearAnswerRows();
        List<String> predictList = getPredictList();
        List<Map<String, Row>> secondReportRowsList = new ArrayList<>();

        for (String predict : predictList) {
            List<Map<String, Row>> rowsWithAnswer = filterListForPredict(rowsWithoutClearAnswersList, predict);

            if (rowsWithAnswer.size() >= 20) {
                List<Map<String, Row>> uniqueRowsCategoryList = createUniqueRowsCategoryList(firstReportRowsList, rowsWithAnswer);
                List<Map<String, Row>> shuffledTenRows = getShuffledTenRows(uniqueRowsCategoryList);
                secondReportRowsList.addAll(shuffledTenRows);
            } else if (rowsWithAnswer.size() >= 10) {
                List<Map<String, Row>> uniqueRowsCategoryList = createUniqueRowsCategoryList(firstReportRowsList, rowsWithAnswer);
                if (uniqueRowsCategoryList.size() < 10) {
                    secondReportRowsList.addAll(uniqueRowsCategoryList);
                } else {
                    List<Map<String, Row>> shuffledTenRows = getShuffledTenRows(rowsWithAnswer);
                    secondReportRowsList.addAll(shuffledTenRows);
                }
            } else {
                secondReportRowsList.addAll(rowsWithAnswer);
            }
        }
        createReport(workBook, sheet, secondReportRowsList);
    }

    private static List<Map<String, Row>> getShuffledTenRows(List<Map<String, Row>> rowsWithoutAnswerList) {
        Collections.shuffle(rowsWithoutAnswerList);
        List<Map<String, Row>> categoryList = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            categoryList.add(rowsWithoutAnswerList.get(i));
        }
        return categoryList;
    }

    private static List<Map<String, Row>> createUniqueRowsCategoryList(List<Map<String, Row>> firstReportRowsList, List<Map<String, Row>> rowsWithoutAnswerList) {
        List<Map<String, Row>> uniqueRowsCategoryList = new ArrayList<>();
        for (int i = 0; i < rowsWithoutAnswerList.size(); i++) {
            for (Map<String, Row> firstReportRowMap : firstReportRowsList) {
                if (!rowsWithoutAnswerList.get(i).entrySet().containsAll(firstReportRowMap.entrySet())) {
                    Map<String, Row> tempMap = new HashMap<>(rowsWithoutAnswerList.get(i));
                    uniqueRowsCategoryList.add(tempMap);
                }
            }
        }
        return uniqueRowsCategoryList;
    }

    private static List<Map<String, Row>> filterClearAnswerRows() {
        List<Map<String, Row>> rowsWithoutClearAnswerList = new ArrayList<>();
        XSSFWorkbook workbook = getWorkBook();
        Sheet dataSheet = workbook.getSheetAt(DATA_SHEET_INDEX);

        Iterator<Row> rowIterator = dataSheet.iterator();
        Row headersRow = rowIterator.next();
        Map<String, Row> headersMap = createHeadersMap(headersRow);
        rowsWithoutClearAnswerList.add(headersMap);
        while (rowIterator.hasNext()) {
            Row currentRow = rowIterator.next();
            Cell answerCell = currentRow.getCell(ANSWER_CELL_COLUMN_INDEX);
            if (answerCell != null) {
                String predictCellValue = getPredictCellValue(currentRow);
                Map<String, Row> rowMap = new HashMap<>();
                rowMap.put(predictCellValue, currentRow);
                rowsWithoutClearAnswerList.add(rowMap);
            }
        }
        return rowsWithoutClearAnswerList;
    }

    static Map<String, Row> createHeadersMap(Row headersRow) {
        Map<String, Row> headersMap = new HashMap<>();
        headersMap.put("headers", headersRow);
        return headersMap;
    }
}
