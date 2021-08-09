package excel_parser;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.*;

import static excel_parser.ExcelReader.PREDICT_CELL_COLUMN_INDEX;
import static excel_parser.ExcelWriter.ANSWER_CELL_COLUMN_INDEX;
import static excel_parser.ExcelWriter.createReport;

public class ExcelReportsUtil {

    private static List<Map<String, Row>> rowsWithoutAnswerList = new ArrayList<>();

    public static List<Map<String, Row>> getRowsWithoutAnswerList() {
        return rowsWithoutAnswerList;
    }

    public static void createFirstReport(XSSFWorkbook workBook) {
        List<Map<String, Row>> filteredRowsList = FirstExcelReport.getFirstReportRowsList();
        XSSFSheet sheet = workBook.createSheet("—рез ≈ва");
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
        rowsWithoutAnswerList.addAll(rowsListWithoutAnswer);
        return rowsListWithoutAnswer;
    }

    private static Map<String, Row> createNamedRowMap(String predictNumOrName, Row row) {
        Map<String, Row> predictRowMap = new HashMap<>();
        predictRowMap.put(predictNumOrName, row);
        return predictRowMap;
    }

    public static String getPredictCellValue(Row row) {
        Cell predictCell = row.getCell(PREDICT_CELL_COLUMN_INDEX);
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

    public static List<Map<String, Row>> createUniqueRowsList(List<Map<String, Row>> oldReportRowsList, List<Map<String, Row>> newReportRowsList) {
        List<Map<String, Row>> uniqueRowsCategoryList = new ArrayList<>();
        for (int i = 0; i < newReportRowsList.size(); i++) {
            for (Map<String, Row> oldReportRowMap : oldReportRowsList) {
                Map<String, Row> newReportRowMap = newReportRowsList.get(i);
                if (!newReportRowMap.entrySet().containsAll(oldReportRowMap.entrySet())) {
                    Map<String, Row> uniqueRowMap = new HashMap<>(newReportRowMap);
                    uniqueRowsCategoryList.add(uniqueRowMap);
                }
            }
        }
        return uniqueRowsCategoryList;
    }

}
