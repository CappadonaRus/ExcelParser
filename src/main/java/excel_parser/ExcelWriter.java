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
import static excel_parser.FirstExcelReport.createFirstReport;
import static excel_parser.SecondExcelReport.createSecondReport;

public class ExcelWriter {

    private static final int CATEGORY_CELL_COUNT = 19;
    private static final int STATISTIC_CELL_NUMBER = CATEGORY_CELL_COUNT + 1;
    private static final int CATEGORY_NAME_COLUMN_INDEX = 0;
    public static final int ANSWER_CELL_COLUMN_INDEX = 4;
    public static final int DATA_SHEET_INDEX = 0;

    public static void writeToExcel() {
        XSSFWorkbook workBook = getWorkBook();
        createFirstReport(workBook);
        createSecondReport(workBook);
        writeSheetIntoBook(workBook);
    }


    static void createReport(XSSFWorkbook workBook, XSSFSheet sheet, List<Map<String, Row>> predictList) {
        String lastCategory = "";
        int rowCount = 0;
        for (Map<String, Row> predictMap : predictList) {
            for (Map.Entry<String, Row> rowMap : predictMap.entrySet()) {
                String rowCategory = rowMap.getKey();
                if (!lastCategory.equals(rowCategory) && isHeadersRow(rowCategory)) {
                    XSSFRow categoryRow = sheet.createRow(rowCount);
                    createCategoryRow(categoryRow, workBook, rowCategory);
                    lastCategory = rowCategory;
                    rowCount++;
                }
                XSSFRow rowForSave = sheet.createRow(rowCount);
                Row predictRow = rowMap.getValue();
                savePredictRowCells(rowForSave, predictRow);
            }
            rowCount++;
        }
    }

    private static void savePredictRowCells(XSSFRow row, Row predictRow) {
        int numberOfCells = predictRow.getPhysicalNumberOfCells();
        for (int j = 0; j < numberOfCells; j++) {
            getCellContentViaTypeAndWrite(row.createCell(j), predictRow.getCell(j));
        }
    }

    private static void writeSheetIntoBook(XSSFWorkbook workBook) {
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

    private static void getCellContentViaTypeAndWrite(Cell createdCell, Cell predictCell) {
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
            DataFormatter nullTypeFormatter = new DataFormatter();
            createdCell.setCellValue(nullTypeFormatter.formatCellValue(predictCell));
        }
    }
}
