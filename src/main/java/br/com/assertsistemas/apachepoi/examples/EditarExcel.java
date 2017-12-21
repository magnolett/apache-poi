package br.com.assertsistemas.apachepoi.examples;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class EditarExcel {

	private static final String fileName = "C:/Users/Marcos/Desktop/novo.xls";

	public static void main(String[] args) {
		
		try {
			FileInputStream file = new FileInputStream (new File(EditarExcel.fileName));
			
			HSSFWorkbook workbook = new HSSFWorkbook(file);
			HSSFSheet sheetAlunos = workbook.getSheetAt(0);
			
			for (int i = 0; i < sheetAlunos.getPhysicalNumberOfRows(); i++) {
				Row row = sheetAlunos.getRow(i);
				Cell cellNota1 = row.getCell(2);
				if (cellNota1.getNumericCellValue() < 9) {
					cellNota1.setCellValue(cellNota1.getNumericCellValue() + 1);
				} else {
					cellNota1.setCellValue(10);
				}
				
				Cell cellNota2 = row.getCell(3);
				Cell cellMedia = row.getCell(4);
				cellMedia.setCellValue((cellNota1.getNumericCellValue() + cellNota2.getNumericCellValue()) / 2);
			}
			file.close();
			
			FileOutputStream outFile = new FileOutputStream(new File(EditarExcel.fileName));
			workbook.write(outFile);
			outFile.close();
			System.out.println("Arquivo excel editado com sucesso!");
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			System.out.println("Arquivo excel não encontrado!");
		} catch (IOException e) {
			e.printStackTrace();	
			System.out.println("Erro na edição do arquivo!");
		}
	}
}
