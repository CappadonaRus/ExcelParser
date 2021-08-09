package excel_parser;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import static excel_parser.Main.getFilePath;

public class ExcelWriter {

    private static final int CATEGORY_CELL_COUNT = 19;
    private static final int STATISTIC_CELL_NUMBER = CATEGORY_CELL_COUNT + 1;
    private static final int CATEGORY_NAME_COLUMN_INDEX = 0;
    public static final int ANSWER_CELL_COLUMN_INDEX = 4;
    public static final int DATA_SHEET_INDEX = 0;

    private static List<String> categoriesList = new ArrayList<>();

    static {
        categoriesList.add("headers");
        for (int i = 1; i <= 125; i++) {
            categoriesList.add(String.valueOf(i));
        }
    }

    public static List<String> getCategoriesList() {
        return categoriesList;
    }

    static void createReport(XSSFWorkbook workBook, XSSFSheet sheet, List<Map<String, Row>> reportRowsList) {
        String lastCategory = "";
        int rowCount = 0;
        for (Map<String, Row> predictMap : reportRowsList) {
            for (Map.Entry<String, Row> rowMap : predictMap.entrySet()) {
                String rowCategory = rowMap.getKey();
                if (!lastCategory.equals(rowCategory) && ExcelReportsUtil.isHeadersRow(rowCategory)) {
                    XSSFRow categoryRow = sheet.createRow(rowCount);
                    createCategoryRow(categoryRow, workBook, rowCategory);
                    lastCategory = rowCategory;
                    rowCount++;
                }
                XSSFRow rowForSave = sheet.createRow(rowCount);
                Row predictRow = rowMap.getValue();
                copyCellsForNewRow(rowForSave, predictRow);
            }
            rowCount++;
        }
    }

    private static void copyCellsForNewRow(XSSFRow newRow, Row row) {
        int numberOfCells = row.getPhysicalNumberOfCells();
        for (int j = 0; j < numberOfCells; j++) {
            writeCellContentViaType(newRow.createCell(j), row.getCell(j));
        }
    }

    public static void writeSheetIntoBook(XSSFWorkbook workBook) {
        try {
            String filePath = getFilePath();
            FileOutputStream fileOut = new FileOutputStream(filePath);
            workBook.write(fileOut);
            fileOut.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void createCategoryRow(XSSFRow categoryRow, XSSFWorkbook workbook, String rowCategory) {
        String categoryName = "Категория: " + rowCategory;
        CellStyle categoryStyle = createCellsStyle(workbook, IndexedColors.YELLOW);
        categoryRow.createCell(CATEGORY_NAME_COLUMN_INDEX).setCellValue(categoryName);
        categoryRow.getCell(CATEGORY_NAME_COLUMN_INDEX).setCellStyle(categoryStyle);
        createCategoryCells(categoryStyle, categoryRow);
        createStatisticCell(workbook, categoryRow);
    }

    private static void createStatisticCell(XSSFWorkbook workbook, XSSFRow categoryRow) {
        CellStyle statisticCell = createCellsStyle(workbook, IndexedColors.RED);
        categoryRow.createCell(STATISTIC_CELL_NUMBER).setCellStyle(statisticCell);
    }

    private static void createCategoryCells(CellStyle categoryStyle, XSSFRow categoryRow) {
        for (int i = 1; i <= CATEGORY_CELL_COUNT; i++) {
            categoryRow.createCell(i);
            categoryRow.getCell(i).setCellStyle(categoryStyle);
        }
    }

    private static CellStyle createCellsStyle(XSSFWorkbook workbook, IndexedColors color) {
        CellStyle categoryStyle = workbook.createCellStyle();
        categoryStyle.setFillForegroundColor(color.getIndex());
        categoryStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return categoryStyle;
    }

    private static void writeCellContentViaType(Cell newCell, Cell cell) {
        try {
            switch (cell.getCellType()) {
                case BOOLEAN:
                    newCell.setCellValue(cell.getBooleanCellValue());
                    break;
                case NUMERIC:
                    newCell.setCellValue(cell.getNumericCellValue());
                    break;
                case STRING:
                    newCell.setCellValue(cell.getStringCellValue());
                    break;
            }
        } catch (NullPointerException npe) {
            // if null set null
            DataFormatter nullTypeFormatter = new DataFormatter();
            newCell.setCellValue(nullTypeFormatter.formatCellValue(cell));
        }
    }
}
