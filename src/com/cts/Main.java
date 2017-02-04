package com.cts;

import com.cts.exceloperation.ReaderEx;
import com.cts.exceloperation.WriterEx;

import java.io.IOException;
import java.util.Date;

public class Main {

    public static void main(String[] args) throws IOException {
//        ReaderEx readerEx = new ReaderEx();
        WriterEx writerEx= new WriterEx("firstsheet");
        writerEx.writeTextData("1. value", 0, 0);
        writerEx.writeDateFormatDataToCell(new Date(),0,1);
        writerEx.wirteToExcelAndClose();

    }


}
