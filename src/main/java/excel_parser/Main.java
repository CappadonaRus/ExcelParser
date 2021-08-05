package excel_parser;

import static excel_parser.ExcelReader.getExcelDate;
import static excel_parser.ExcelWriter.writeToExcel;
import static excel_parser.ExcelWriterUtil.filterRows;

public class Main {

    public static void main(String[] args) {
        getExcelDate();
        filterRows();
        writeToExcel();
    }
}
