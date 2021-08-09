package excel_parser;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

import static excel_parser.ExcelWriter.ANSWER_CELL_COLUMN_INDEX;

public class ExcelReader {

    public static final int PREDICT_CELL_COLUMN_INDEX = 5;
    private static String filePath = "./—рез ≈ва 09.08.xlsx";
    private static List<Map<String, Row>> firstReportRowsList = new ArrayList<>();
    private static XSSFWorkbook workbook;

    public static List<Map<String, Row>> getFirstReportRowsList() {
        return firstReportRowsList;
    }

    public static XSSFWorkbook getWorkBook() {
        return workbook;
    }

    public static String getFilePath() {
        return filePath;
    }

    public static void readExcelAndGetSheetData() {
        readExcelWorkBook(filePath);
        Sheet dataSheet = workbook.getSheetAt(0);
        readSheetRowsAndCreatePredictList(dataSheet);
    }


    private static void readExcelWorkBook(String filePath) {
        try {
            FileInputStream inputStream = new FileInputStream(new File(filePath));
            workbook = new XSSFWorkbook(inputStream);
            inputStream.close();
        } catch (IOException fileNotFoundException) {
            fileNotFoundException.printStackTrace();
        }
    }

    private static void readSheetRowsAndCreatePredictList(Sheet dataSheet) {
        Iterator<Row> rowIterator = dataSheet.iterator();
        Row headersRow = rowIterator.next();
        Map<String, Row> headersMap = createPredictRowMap("headers", headersRow);
        firstReportRowsList.add(headersMap);

        while (rowIterator.hasNext()) {
            Row currentRow = rowIterator.next();
            Cell answerCell = currentRow.getCell(ANSWER_CELL_COLUMN_INDEX);
            if (answerCell != null) {
                String predictCategoryNum = getPredictCellValue(currentRow);
                Map<String, Row> predictMap = createPredictRowMap(predictCategoryNum, currentRow);
                firstReportRowsList.add(predictMap);
            }
        }
    }

    private static Map<String, Row> createPredictRowMap(String predictNumber, Row row) {
        Map<String, Row> predictRowMap = new HashMap<>();
        predictRowMap.put(predictNumber, row);
        return predictRowMap;
    }

    public static String getPredictCellValue(Row currentRow) {
        Cell predictCell = currentRow.getCell(PREDICT_CELL_COLUMN_INDEX);
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
}
