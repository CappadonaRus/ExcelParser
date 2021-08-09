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

public class ExcelReader {

    public static final int PREDICT_CELL_COLUMN_INDEX = 5;
    private static String filePath = "./—рез ≈ва 09.08.xlsx";
    private static List<Map<String, Row>> excelRowsMap = new ArrayList<>();
    private static XSSFWorkbook workbook;

    public static List<Map<String, Row>> getExcelRowMap() {
        return excelRowsMap;
    }

    public static XSSFWorkbook getWorkBook() {
        return workbook;
    }

    public static String getFilePath() {
        return filePath;
    }

    public static List<Map<String, Row>> getExcelRowsMap() {
        return excelRowsMap;
    }

    public static void getExcelDate() {

        try {
            FileInputStream inputStream = new FileInputStream(new File(filePath));
            workbook = new XSSFWorkbook(inputStream);
            Sheet dataSheet = workbook.getSheetAt(0);

            readSheetRowsAndCreatePredictList(dataSheet);

            inputStream.close();

        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    private static void readSheetRowsAndCreatePredictList(Sheet dataSheet) {
        Iterator<Row> rowIterator = dataSheet.iterator();
        Row headersRow = rowIterator.next();
        createHeadersMap(headersRow);

        while (rowIterator.hasNext()) {
            Row currentRow = rowIterator.next();
            String predictCellValue = getPredictCellValue(currentRow);
            createPredictRowMap(currentRow, predictCellValue);
        }
    }

    static void createHeadersMap(Row headersRow) {
        Map<String, Row> headersMap = new HashMap<>();
        headersMap.put("headers", headersRow);
        excelRowsMap.add(headersMap);
    }

    public static String getPredictCellValue(Row currentRow) {
        Cell predictCell = currentRow.getCell(PREDICT_CELL_COLUMN_INDEX);
        String predictCellValue = "";
        switch (predictCell.getCellType()) {
            case STRING:
                //skip header
                break;

            default:
                DataFormatter stringFormatter = new DataFormatter();
                predictCellValue = stringFormatter.formatCellValue(predictCell);
                break;
        }
        return predictCellValue;
    }

    private static void createPredictRowMap(Row currentRow, String predictCellValue) {
        Map<String, Row> predictMap = new HashMap<>();
        predictMap.put(predictCellValue, currentRow);
        excelRowsMap.add(predictMap);
    }

}
