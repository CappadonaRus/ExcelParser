package ru.excel;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.List;
import java.util.Map;

public interface Reportable {

    List<Map<String, Row>> createReport(XSSFWorkbook workBook, int sheetIndex);
}
