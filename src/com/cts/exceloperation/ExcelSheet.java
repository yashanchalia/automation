package com.cts.exceloperation;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Created by Raj on 2/4/2017.
 */
public class ExcelSheet {
    HSSFWorkbook workbook;

    public ExcelSheet(){
        workbook = new HSSFWorkbook();
    }

    public void wirteToExcelAndClose() throws IOException {
        workbook.write(new FileOutputStream("firstTry.xls"));
        workbook.close();
    }

}
