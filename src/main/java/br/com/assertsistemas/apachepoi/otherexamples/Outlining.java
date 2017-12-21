package br.com.assertsistemas.apachepoi.otherexamples;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Outlining {

    public static void main(String[] args) throws IOException {
        Outlining o=new Outlining();
        o.groupRowColumn();
        o.collapseExpandRowColumn();
    }


    private void groupRowColumn() throws IOException {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet1 = wb.createSheet("new sheet");

        sheet1.groupRow( 5, 14 );
        sheet1.groupRow( 7, 14 );
        sheet1.groupRow( 16, 19 );

        sheet1.groupColumn( (short)4, (short)7 );
        sheet1.groupColumn( (short)9, (short)12 );
        sheet1.groupColumn( (short)10, (short)11 );

        OutputStream fileOut = new FileOutputStream("C:/Users/Marcos/Desktop/outlining.xlsx");
        try {
            wb.write(fileOut);
        } finally {
            fileOut.close();
            wb.close();
        }
    }

    private void collapseExpandRowColumn() throws IOException {
        Workbook wb2 = new XSSFWorkbook();
        Sheet sheet2 = wb2.createSheet("new sheet");
        sheet2.groupRow( 5, 14 );
        sheet2.groupRow( 7, 14 );
        sheet2.groupRow( 16, 19 );

        sheet2.groupColumn( (short)4, (short)7 );
        sheet2.groupColumn( (short)9, (short)12 );
        sheet2.groupColumn( (short)10, (short)11 );


        sheet2.setRowGroupCollapsed( 7, true );
        //sheet1.setRowGroupCollapsed(7,false);

        sheet2.setColumnGroupCollapsed( (short)4, true );
        sheet2.setColumnGroupCollapsed( (short)4, false );

        OutputStream fileOut = new FileOutputStream("outlining_collapsed.xlsx");
        try {
            wb2.write(fileOut);
        } finally {
            fileOut.close();
            wb2.close();
        }
    }
}