package ExcelParser;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;

import static ExcelParser.ExcelReader.getFilePath;
import static ExcelParser.ExcelReader.getWorkBook;
import static ExcelParser.ExcelWriterUtil.getFilteredMap;

public class ExcelWriter {

    public static void writeToExcel() {
        XSSFWorkbook workBook = getWorkBook();
        XSSFSheet sheet = workBook.createSheet("Auto srez");

        List<Map<String, Row>> filteredMap = getFilteredMap();
        String lastCategory = "";
        int rowCount = 0;
        for (Map<String, Row> filteredRowMap : filteredMap) {
            for (Map.Entry<String, Row> rowMap : filteredRowMap.entrySet()) {
                String rowCategory = rowMap.getKey();
                if (!lastCategory.equals(rowCategory) && !rowCategory.equals("headers")) {
                    createCategoryRow(workBook, sheet, rowCount, rowCategory);
                    lastCategory = rowCategory;
                    rowCount++;
                }
                XSSFRow createdRow = sheet.createRow(rowCount);
                Row row = rowMap.getValue();
                int numberOfCells = row.getPhysicalNumberOfCells();
                for (int j = 0; j < numberOfCells; j++) {
                    String cellValue = getCellValueAsString(row.getCell(j));
                    createdRow.createCell(j).setCellValue(cellValue);
                }
            }
            rowCount++;
        }
        try {
            String filePath = getFilePath();
            OutputStream fileOut = new FileOutputStream(filePath);
            workBook.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void createCategoryRow(XSSFWorkbook workBook, XSSFSheet sheet, int rowCount, String rowCategory) {
        CellStyle style = workBook.createCellStyle();
        style.setFillBackgroundColor(IndexedColors.YELLOW.getIndex());
        style.setFillPattern(CellStyle.BIG_SPOTS);
        XSSFRow categoryRow = sheet.createRow(rowCount);
        categoryRow.setRowStyle(style);
        categoryRow.createCell(1).setCellValue("Category: " + rowCategory);
    }

    private static String getCellValueAsString(Cell predictCell) {
        try {
            switch (predictCell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                    return predictCell.getStringCellValue();

                default:
                    DataFormatter stringFormatter = new DataFormatter();
                    return stringFormatter.formatCellValue(predictCell);
            }
        } catch (NullPointerException npe) {
            // hoho
        }
        return null;
    }
}
