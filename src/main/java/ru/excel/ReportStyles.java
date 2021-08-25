package ru.excel;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReportStyles {

    static void addBoldIdCell(XSSFWorkbook workBook, XSSFRow rowForSave) {
        XSSFCellStyle idBoldStyle = workBook.createCellStyle();
        createBold(workBook, idBoldStyle);
        idBoldStyle.setBorderLeft(BorderStyle.THIN);
        idBoldStyle.setBorderRight(BorderStyle.THIN);
        idBoldStyle.setBorderBottom(BorderStyle.THIN);
        idBoldStyle.setBorderTop(BorderStyle.THIN);
        rowForSave.getCell(0).setCellStyle(idBoldStyle);
        idBoldStyle.setAlignment(HorizontalAlignment.CENTER);
    }

    static XSSFCellStyle createBorderStyle(XSSFWorkbook workBook) {
        XSSFCellStyle cellBorder = workBook.createCellStyle();
        cellBorder.setBorderLeft(BorderStyle.THIN);
        cellBorder.setBorderRight(BorderStyle.THIN);
        cellBorder.setBorderBottom(BorderStyle.THIN);
        cellBorder.setBorderTop(BorderStyle.THIN);
        return cellBorder;
    }

    static void createBold(XSSFWorkbook workbook, CellStyle categoryStyle) {
        XSSFFont fontBold = workbook.createFont();
        fontBold.setBold(true);
        categoryStyle.setFont(fontBold);
    }

}
