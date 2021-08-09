package excel_parser;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

import static excel_parser.ExcelReader.*;
import static excel_parser.ExcelWriterUtil.*;

public class ExcelWriter {

    private static final int CATEGORY_CELL_COUNT = 19;
    private static final int STATISTIC_CELL_NUMBER = CATEGORY_CELL_COUNT + 1;
    private static final int CATEGORY_NAME_COLUMN_INDEX = 0;
    public static final int ANSWER_CELL_COLUMN_INDEX = 4;
    public static final int DATA_SHEET_INDEX = 0;

    public static void writeToExcel() {
        XSSFWorkbook workBook = getWorkBook();
        List<Map<String, Row>> filteredRowsList = getFilteredMap();
        createFirstReport(workBook, filteredRowsList);
        createSecondReport(workBook, filteredRowsList);
        writeSheetIntoBook(workBook);
    }

    private static void createFirstReport(XSSFWorkbook workBook, List<Map<String, Row>> filteredRowsList) {
        XSSFSheet sheet = workBook.createSheet("Срез Ева");
        createReport(workBook, sheet, filteredRowsList);
    }

    private static void createSecondReport(XSSFWorkbook workBook, List<Map<String, Row>> firstReportRowsList) {
        XSSFSheet sheet = workBook.createSheet("Срез Ева без оператора");
        List<Map<String, Row>> rowsWithoutClearAnswersList = filterClearAnswerRows();
        List<String> predictList = getPredictList();
        List<Map<String, Row>> secondReportList = new ArrayList<>();
//        addHeadersRow(rowsWithoutClearAnswersList, secondReportList);

        for (String predict : predictList) {
            List<Map<String, Row>> rowsWithAnswer = filterListForPredict(rowsWithoutClearAnswersList, predict);

            if (rowsWithAnswer.size() >= 20) {
                List<Map<String, Row>> uniqueRowsCategoryList = createUniqueRowsCategoryList(firstReportRowsList, rowsWithAnswer);
                List<Map<String, Row>> shuffledTenRows = getShuffledTenRows(uniqueRowsCategoryList);
                secondReportList.addAll(shuffledTenRows);
            } else if (rowsWithAnswer.size() >= 10) {
                List<Map<String, Row>> uniqueRowsCategoryList = createUniqueRowsCategoryList(firstReportRowsList, rowsWithAnswer);
                if (uniqueRowsCategoryList.size() < 10) {
                    secondReportList.addAll(uniqueRowsCategoryList);
                } else {
                    List<Map<String, Row>> shuffledTenRows = getShuffledTenRows(rowsWithAnswer);
                    secondReportList.addAll(shuffledTenRows);
                }
            } else {
                secondReportList.addAll(rowsWithAnswer);
            }
        }
        createReport(workBook, sheet, secondReportList);
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

    private static void addHeadersRow(List<Map<String, Row>> rowsWithoutClearAnswersList, List<Map<String, Row>> secondReportList) {
        for (Map<String, Row> rowMap : rowsWithoutClearAnswersList) {
            for (Map.Entry<String, Row> row : rowMap.entrySet()) {
                if (row.getKey().equals("headers")) {
                    secondReportList.add(rowMap);
                    break;
                }
            }
        }
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

    private static void createReport(XSSFWorkbook workBook, XSSFSheet sheet, List<Map<String, Row>> predictList) {
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

    private static void writeSheetIntoBook(XSSFWorkbook workBook) {
        try {
            String filePath = getFilePath();
            FileOutputStream fileOut = new FileOutputStream(filePath);
            workBook.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void savePredictRowCells(XSSFRow row, Row predictRow) {
        int numberOfCells = predictRow.getPhysicalNumberOfCells();
        for (int j = 0; j < numberOfCells; j++) {
            getCellContentViaTypeAndWrite(row.createCell(j), predictRow.getCell(j));
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
