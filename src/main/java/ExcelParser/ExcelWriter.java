package ExcelParser;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;

import static ExcelParser.ExcelReader.getFilePath;
import static ExcelParser.ExcelReader.getWorkBook;
import static ExcelParser.ExcelWriterUtil.getFilteredMap;

public class ExcelWriter {

    public static void writeToExcel() {
        XSSFWorkbook workBook = getWorkBook();
        XSSFSheet sheet = workBook.createSheet("Auto srez");

        List<Map<String, Row>> filteredMap = getFilteredMap();
        String lastCategory = "";
        for (int i = 0; i < filteredMap.size(); i++) {
            XSSFRow createdRow = sheet.createRow(i);
            for (Map.Entry<String, Row> rowMap : filteredMap.get(i).entrySet()) {
                String category = rowMap.getKey();
//                if(!lastCategory.equals(category)){
//
//                }
                Row row = rowMap.getValue();
                int numberOfCells = row.getPhysicalNumberOfCells();
                for (int j = 0; j < numberOfCells; j++) {
                    String cellValue = getCellValueAsString(row.getCell(j));
                    createdRow.createCell(j).setCellValue(cellValue);
                }
            }
        }
        try {
            String filePath = getFilePath();
            OutputStream fileOut = new FileOutputStream(filePath);
            workBook.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static String getCellValueAsString(Cell predictCell) {
        try {
            switch (predictCell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                    return predictCell.getStringCellValue();

                default:
                    DataFormatter stringFormatter = new DataFormatter();
                    return stringFormatter.formatCellValue(predictCell);
            }
        } catch (NullPointerException npe) {
            // hoho
        }
        return null;
    }
}
