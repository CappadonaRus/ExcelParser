package ExcelParser;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

public class ExcelReader {

    private static String filePath = "./eva.xlsx";

    private static List<Map<String, Row>> excelRowsMap = new ArrayList<>();
    private static XSSFWorkbook workbook;

    public static List<Map<String, Row>> getExcelRowMap() {
        return excelRowsMap;
    }

    public static XSSFWorkbook getWorkBook(){
        return workbook;
    }

    public static String getFilePath(){
        return filePath;
    }

    public static void getExcelDate() {

        try {
            FileInputStream inputStream = new FileInputStream(new File(filePath));
            workbook = new XSSFWorkbook(inputStream);
            Sheet datatypeSheet = workbook.getSheetAt(0);

            Iterator<Row> iterator = datatypeSheet.iterator();

            while (iterator.hasNext()) {
                Row currentRow = iterator.next();
                Cell predictCell = currentRow.getCell(5);
                getPredictCellValue(currentRow, predictCell);
            }

            inputStream.close();

        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    private static void getPredictCellValue(Row currentRow, Cell predictCell) {
        switch (predictCell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                //skip header
                break;

            default:
                DataFormatter stringFormatter = new DataFormatter();
                String predictCellValue = stringFormatter.formatCellValue(predictCell);
                createPredictRowMap(currentRow, predictCellValue);
                break;

        }
    }

    private static void createPredictRowMap(Row currentRow, String predictCellValue) {
        Map<String, Row> predictMap = new HashMap<>();
        predictMap.put(predictCellValue, currentRow);
        excelRowsMap.add(predictMap);
    }

}
