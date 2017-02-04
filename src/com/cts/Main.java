package com.cts;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

public class Main {

    public static void main(String[] args) throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        DataFormat dataFormat = workbook.createDataFormat();

        //create worksheet
        HSSFSheet sheet = workbook.createSheet("firstsheet");

        //add general string to cell (0,0)
        HSSFRow row = sheet.createRow(0);
        row.createCell(0).setCellValue("1. value");

        //add style to particular cell
        HSSFCell cell2 = row.createCell(1);
        CellStyle dateStyle = getDateStyle(workbook, dataFormat);
        cell2.setCellStyle(dateStyle);
        cell2.setCellValue(new Date());

        sheet.autoSizeColumn(1);
        workbook.write(new FileOutputStream("firstTry.xls"));
        workbook.close();
    }

    private static CellStyle getDateStyle(HSSFWorkbook workbook, DataFormat dataFormat) {
        CellStyle dateStyle = workbook.createCellStyle();
        dateStyle.setDataFormat(dataFormat.getFormat("dd.mm.yyyy"));
        return dateStyle;
    }
}
