package ExcelParser;

import static ExcelParser.ExcelReader.getExcelDate;
import static ExcelParser.ExcelWriter.writeToExcel;
import static ExcelParser.ExcelWriterUtil.filterRows;

public class Main {

    public static void main(String[] args) {
        getExcelDate();
        filterRows();
        writeToExcel();
    }
}
