package ru.excel.late_report;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.*;
import ru.excel.ExcelReader;
import ru.excel.late_report.model.CategoryRow;
import ru.excel.late_report.model.ReportRow;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;

import static ru.excel.late_report.util.LateReportUtil.createCategoriesMap;
import static ru.excel.late_report.util.LateReportUtil.createReportMap;


public class LateReportMain {

    public static final int FIRST_SHEET_WORKBOOK_INDEX = 0;

    private static String filePath = "./discharge.xlsx";
    private static String categoriesFilePath = "./category.xlsx";
    private static String dayFilePath = "./day.xlsx";

    public static void main(String[] args) {
        filePath = args[0] + ".xlsx";
        var categories = ExcelReader.readExcelWorkBook(categoriesFilePath);
        var allCategories = createCategoriesMap(categories.getSheetAt(FIRST_SHEET_WORKBOOK_INDEX));
        var dayCategoriesBook = ExcelReader.readExcelWorkBook(dayFilePath);
        var dayCategories = createCategoriesMap(
                dayCategoriesBook.getSheetAt(FIRST_SHEET_WORKBOOK_INDEX),
                2);

        var searchCategories = filterCategories(allCategories, dayCategories);

        var report = ExcelReader.readExcelWorkBook(filePath);
        var reportMap = createReportMap(report.getSheetAt(FIRST_SHEET_WORKBOOK_INDEX));

        var sortedCategoriesRows = sortRowsByCategories(reportMap, searchCategories);
        var sortedMap = sortMap(sortedCategoriesRows);

        var sortedCountryMap = sortCategoryByCountry(sortedMap);

        var finalReport = createReportRows(sortedCountryMap);

        createFinalReport(searchCategories, finalReport);

    }

    private static LinkedHashMap<String, List<CategoryRow>> sortMap(Map<String, List<CategoryRow>> sortedCategoriesRows) {
        return sortedCategoriesRows
                .entrySet()
                .stream()
                .sorted(Collections.reverseOrder(Comparator.comparing(l -> l.getValue().size())))
                .collect(Collectors.toMap(Map.Entry::getKey, Map.Entry::getValue, (v1, v2) -> v1, LinkedHashMap::new));
    }

    private static void createFinalReport(HashMap<String, String> searchCategories, LinkedHashMap<String, ReportRow> finalReport) {
        var finalRep = new XSSFWorkbook();
        var reportSheet = finalRep.createSheet("отчет");
        createLateReportHeaders(reportSheet);
        AtomicInteger count = new AtomicInteger();
        count.getAndIncrement();
        finalReport.forEach((category, report) -> {
            var row = reportSheet.createRow(count.get());
            row.createCell(0).setCellValue(searchCategories.get(category));
            row.createCell(1).setCellValue(Integer.parseInt(report.getRuEva()));
            row.createCell(2).setCellValue(Integer.parseInt(report.getRuOperator()));
            row.createCell(3).setCellValue(Integer.parseInt(report.getByEva()));
            row.createCell(4).setCellValue(Integer.parseInt(report.getByOperator()));
            row.createCell(5).setCellValue(Integer.parseInt(report.getKzEva()));
            row.createCell(6).setCellValue(Integer.parseInt(report.getKzOperator()));
            row.createCell(7).setCellValue(Integer.parseInt(report.getKgEva()));
            row.createCell(8).setCellValue(Integer.parseInt(report.getKgOperator()));
            row.createCell(9).setCellValue(Integer.parseInt(report.getAmEva()));
            row.createCell(10).setCellValue(Integer.parseInt(report.getAmOperator()));
            row.createCell(11).setCellValue(Integer.parseInt(report.getUzEva()));
            row.createCell(12).setCellValue(Integer.parseInt(report.getUzOperator()));
            count.getAndIncrement();
        });
        for (int i = 0; i < 13; i++) {
            reportSheet.autoSizeColumn(i);
        }

        try {
            var fos = new FileOutputStream("final_report.xlsx");
            finalRep.write(fos);
            fos.close();
            finalRep.close();
        } catch (IOException e) {
            e.printStackTrace();
        }


    }

    private static void createLateReportHeaders(XSSFSheet reportSheet) {
        var headersRow = reportSheet.createRow(0);
        headersRow.createCell(0).setCellValue("Тема");
        headersRow.createCell(1).setCellValue("E");
        headersRow.createCell(2).setCellValue("O");
        headersRow.createCell(3).setCellValue("E");
        headersRow.createCell(4).setCellValue("O");
        headersRow.createCell(5).setCellValue("E");
        headersRow.createCell(6).setCellValue("O");
        headersRow.createCell(7).setCellValue("E");
        headersRow.createCell(8).setCellValue("O");
        headersRow.createCell(9).setCellValue("E");
        headersRow.createCell(10).setCellValue("O");
        headersRow.createCell(11).setCellValue("E");
        headersRow.createCell(12).setCellValue("O");
        addHeaderBoards(reportSheet.getWorkbook(), headersRow);
    }

    public static void addHeaderBoards(XSSFWorkbook workBook, XSSFRow row) {
        XSSFCellStyle idBoldStyle = workBook.createCellStyle();
        createBold(workBook, idBoldStyle);
        idBoldStyle.setBorderLeft(BorderStyle.THIN);
        idBoldStyle.setBorderRight(BorderStyle.THIN);
        idBoldStyle.setBorderBottom(BorderStyle.THIN);
        idBoldStyle.setBorderTop(BorderStyle.THIN);
        idBoldStyle.setAlignment(HorizontalAlignment.CENTER);
        for (int i = 0; i < 13; i++) {
            row.getCell(i).setCellStyle(idBoldStyle);
        }
    }

    public static void createBold(XSSFWorkbook workbook, CellStyle categoryStyle) {
        XSSFFont fontBold = workbook.createFont();
        fontBold.setBold(true);
        categoryStyle.setFont(fontBold);
    }

    private static HashMap<String, String> filterCategories(Map<String, String> allCategories, ArrayList<String> dayCategories) {
        var filteredCategories = new HashMap<String, String>();
        dayCategories.forEach(dayCategoryName -> {
            var category = allCategories
                    .entrySet()
                    .stream()
                    .filter(key -> key.getValue().trim().contains(dayCategoryName.trim()))
                    .findFirst();

            category.ifPresent(dayCat -> filteredCategories.put(dayCat.getKey(), dayCat.getValue()));
        });

        return filteredCategories;
    }

    private static LinkedHashMap<String, ReportRow> createReportRows(LinkedHashMap<String, Map<String, List<CategoryRow>>> categoriesMap) {
        var sumReport = new LinkedHashMap<String, ReportRow>();

        categoriesMap.forEach((category, countriesRows) -> {
            var reportRow = new ReportRow();
            reportRow.setCategoryName(category);

            countriesRows.forEach((country, countryList) -> {
                var evaSize = getEvaOrOperatorSize(countryList, "1");
                var operatorSize = getEvaOrOperatorSize(countryList, "0");
                setEvaAndOperatorSizeByCountry(reportRow, country, evaSize, operatorSize);
            });
            sumReport.put(category, reportRow);
        });

        return sumReport;
    }

    private static void setEvaAndOperatorSizeByCountry(ReportRow reportRow,
                                                       String country,
                                                       String evaSize,
                                                       String operatorSize) {
        switch (country) {
            case "RU":
                reportRow.setRuEva(evaSize);
                reportRow.setRuOperator(operatorSize);
                break;
            case "BY":
                reportRow.setByEva(evaSize);
                reportRow.setByOperator(operatorSize);
                break;
            case "KZ":
                reportRow.setKzEva(evaSize);
                reportRow.setKzOperator(operatorSize);
                break;
            case "KG":
                reportRow.setKgEva(evaSize);
                reportRow.setKgOperator(operatorSize);
                break;
            case "AM":
                reportRow.setAmEva(evaSize);
                reportRow.setAmOperator(operatorSize);
                break;
            case "UZ":
                reportRow.setUzEva(evaSize);
                reportRow.setUzOperator(operatorSize);
                break;
        }
    }

    private static String getEvaOrOperatorSize(List<CategoryRow> countryList, String evaOrOperatorFlag) {
        return String.valueOf((int) countryList
                .stream()
                .filter(c -> c.getIsEvaAnswered().equals(evaOrOperatorFlag))
                .count());
    }

    private static LinkedHashMap<String, Map<String, List<CategoryRow>>> sortCategoryByCountry(LinkedHashMap<String, List<CategoryRow>> categoryList) {
        var countryMap = new LinkedHashMap<String, Map<String, List<CategoryRow>>>();
        categoryList.forEach((category, categoryRows) -> {
            var tempMap = new HashMap<String, List<CategoryRow>>();
            tempMap.put("RU", sortCategoriesByCountry(categoryRows, "RU"));
            tempMap.put("BY", sortCategoriesByCountry(categoryRows, "BY"));
            tempMap.put("KZ", sortCategoriesByCountry(categoryRows, "KZ"));
            tempMap.put("KG", sortCategoriesByCountry(categoryRows, "KG"));
            tempMap.put("AM", sortCategoriesByCountry(categoryRows, "AM"));
            tempMap.put("UZ", sortCategoriesByCountry(categoryRows, "UZ"));
            countryMap.put(category, tempMap);
        });
        return countryMap;
    }


    private static List<CategoryRow> sortCategoriesByCountry(List<CategoryRow> categoryRows, String countryName) {
        return categoryRows
                .stream()
                .filter(c -> c.getCountry().equals(countryName))
                .collect(Collectors.toList());
    }

    private static Map<String, List<CategoryRow>> sortRowsByCategories(List<CategoryRow> reportMap, Map<String, String> reportCategories) {
        var categoryList = new HashMap<String, List<CategoryRow>>();
        reportCategories.forEach((category, value) ->
                categoryList.put(category,
                        reportMap.stream()
                                .filter(c -> c.getCategory().equals(category))
                                .collect(Collectors.toList())));
        return categoryList;
    }


}
