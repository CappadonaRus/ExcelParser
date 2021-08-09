package excel_parser;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.List;
import java.util.Map;

import static excel_parser.ExcelWriter.createReport;
import static excel_parser.ExcelWriterUtil.getFilteredMap;

public class FirstExcelReport {

    public static void createFirstReport(XSSFWorkbook workBook) {
        List<Map<String, Row>> filteredRowsList = getFilteredMap();
        XSSFSheet sheet = workBook.createSheet("—рез ≈ва");
        createReport(workBook, sheet, filteredRowsList);
    }


}
