package ru.excel.late_report.util;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import ru.excel.late_report.model.CategoryRow;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class LateReportUtil {

    private static final int CATEGORIES_NAME_CELL_INDEX = 0;
    private static final int CATEGORIES_CELL_INDEX = 1;

    private static final int REPORT_CATEGORY_CELL_INDEX = 5;
    private static final int REPORT_IS_EVA_ANSWERED_CELL_INDEX = 8;
    private static final int REPORT_COUNTRY_CELL_INDEX = 14;


    public static Map<String, String> createCategoriesMap(XSSFSheet firstReportSheet) {
        var categoriesMap = new HashMap<String, String>();
        var sheetIterator = firstReportSheet.iterator();

        while (sheetIterator.hasNext()) {
            var currentRow = sheetIterator.next();
            var categoryName = currentRow.getCell(CATEGORIES_NAME_CELL_INDEX);
            var category = currentRow.getCell(CATEGORIES_CELL_INDEX);
            categoriesMap.put(String.valueOf(category.getNumericCellValue()), categoryName.getStringCellValue());
        }
        return categoriesMap;
    }

    public static ArrayList<String> createCategoriesMap(XSSFSheet firstReportSheet, int categoryCellNum) {
        var categoriesMap = new ArrayList<String>();
        var sheetIterator = firstReportSheet.iterator();
        sheetIterator.next();

        while (sheetIterator.hasNext()) {
            var currentRow = sheetIterator.next();
            var categoryName = currentRow.getCell(categoryCellNum);
            categoriesMap.add(categoryName.getStringCellValue());
        }
        return categoriesMap;
    }

    public static List<CategoryRow> createReportMap(XSSFSheet reportSheet) {
        var reportList = new ArrayList<CategoryRow>();
        var reportIterator = reportSheet.iterator();
        reportIterator.next();

        while (reportIterator.hasNext()) {
            var currentRow = reportIterator.next();
            var categoryCell = currentRow.getCell(REPORT_CATEGORY_CELL_INDEX);
            var isEvaAnsweredCell = currentRow.getCell(REPORT_IS_EVA_ANSWERED_CELL_INDEX);
            var countryCell = currentRow.getCell(REPORT_COUNTRY_CELL_INDEX);
            reportList.add(new CategoryRow(
                    String.valueOf(categoryCell.getNumericCellValue()),
                    isEvaAnsweredCell.getStringCellValue(),
                    countryCell.getStringCellValue()));
        }

        return reportList;
    }
}
