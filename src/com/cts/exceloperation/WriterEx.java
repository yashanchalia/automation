package com.cts.exceloperation;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

/**
 * Created by Raj on 2/4/2017.
 */
public class WriterEx extends ExcelSheet{
    HSSFSheet sheet;
    HSSFRow row;

    public WriterEx(String sheetName){
        this.sheet = workbook.createSheet(sheetName);
        row = sheet.createRow(0);
    }

    public void writeTextData(String value, int rowNumber, int columnNumber) throws IOException {
//        HSSFRow row = sheet.createRow(rowNumber);
        row.createCell(columnNumber).setCellValue(value);
        sheet.autoSizeColumn(columnNumber);
    }

    public void writeDateFormatDataToCell(Date dateValue, int rowNumber, int columnNumber) throws IOException {
        DataFormat dataFormat = workbook.createDataFormat();
//        HSSFRow row = sheet.createRow(rowNumber);
        HSSFCell cell2 = row.createCell(1);
        CellStyle dateStyle = getDateStyle(workbook, dataFormat);
        cell2.setCellStyle(dateStyle);
        cell2.setCellValue(dateValue);
        sheet.autoSizeColumn(columnNumber);
    }

    private static CellStyle getDateStyle(HSSFWorkbook workbook, DataFormat dataFormat) {
        CellStyle dateStyle = workbook.createCellStyle();
        dateStyle.setDataFormat(dataFormat.getFormat("dd.mm.yyyy"));
        return dateStyle;
    }
}
