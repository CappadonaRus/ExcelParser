package excel_parser;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;

import static excel_parser.ExcelReader.getFilePath;
import static excel_parser.ExcelReader.getWorkBook;
import static excel_parser.ExcelWriterUtil.getFilteredMap;

public class ExcelWriter {

    public static void writeToExcel() {
        XSSFWorkbook workBook = getWorkBook();
        XSSFSheet sheet = workBook.createSheet("Авто срез");
        List<Map<String, Row>> filteredMap = getFilteredMap();
        String lastCategory = "";
        int rowCount = 0;
        for (Map<String, Row> filteredRowMap : filteredMap) {
            for (Map.Entry<String, Row> rowMap : filteredRowMap.entrySet()) {
                String rowCategory = rowMap.getKey();
                if (!lastCategory.equals(rowCategory) && isHeadersRow(rowCategory)) {
                    createCategoryRow(workBook, sheet, rowCount, rowCategory);
                    lastCategory = rowCategory;
                    rowCount++;
                }
                XSSFRow createdRow = sheet.createRow(rowCount);
                Row row = rowMap.getValue();
                int numberOfCells = row.getPhysicalNumberOfCells();
                for (int j = 0; j < numberOfCells; j++) {
                    getCellValueAsString(createdRow.createCell(j), row.getCell(j));
                }
            }
            rowCount++;
        }
        try {
            String filePath = getFilePath();
            FileOutputStream fileOut = new FileOutputStream(filePath);
            workBook.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static boolean isHeadersRow(String rowCategory) {
        return !rowCategory.equals("headers");
    }

    private static void createCategoryRow(XSSFWorkbook workBook, XSSFSheet sheet, int rowCount, String rowCategory) {
        int cellsCount = 19;
        CellStyle categoryStyle = workBook.createCellStyle();
        categoryStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        categoryStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        XSSFRow categoryRow = sheet.createRow(rowCount);
        String category = "Категория: " + rowCategory;
        categoryRow.createCell(0).setCellValue(category);
        categoryRow.getCell(0).setCellStyle(categoryStyle);
        for (int i = 1; i <= cellsCount; i++) {
            categoryRow.createCell(i);
            categoryRow.getCell(i).setCellStyle(categoryStyle);
        }
        CellStyle kpiCell = workBook.createCellStyle();
        kpiCell.setFillForegroundColor(IndexedColors.RED.getIndex());
        kpiCell.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        categoryRow.createCell(cellsCount + 1).setCellStyle(kpiCell);
    }

    private static void getCellValueAsString(Cell createdCell, Cell predictCell) {
        try {
            switch (predictCell.getCellType()) {
                case BOOLEAN:
                    createdCell.setCellValue(predictCell.getBooleanCellValue());
                    break;
                case NUMERIC:
                    createdCell.setCellValue(predictCell.getNumericCellValue());
                    break;
                case STRING:
                    createdCell.setCellValue(predictCell.getStringCellValue());
                    break;
            }
        } catch (NullPointerException npe) {
            // if null set null
            DataFormatter stringFormatter = new DataFormatter();
            createdCell.setCellValue(stringFormatter.formatCellValue(predictCell));
        }
    }

//    private void resizeColumns(){
//        // Resize all columns to fit the content size
//        for(int i = 0; i < columns.length; i++) {
//            sheet.autoSizeColumn(i);
//        }
//    }

}
