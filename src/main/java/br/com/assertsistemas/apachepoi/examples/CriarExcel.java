package br.com.assertsistemas.apachepoi.examples;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.util.CellRangeAddress;

public class CriarExcel {

	private static final String fileName = "C:/Users/Marcos/Desktop/t.xls";

	public static void main(String[] args) throws IOException {

		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheetAlunos = workbook.createSheet("Alunos");
		HSSFRow row = sheetAlunos.createRow((short) 0);
		CreationHelper factory = workbook.getCreationHelper();
		Drawing drawing = sheetAlunos.createDrawingPatriarch();

		ClientAnchor anchor = factory.createClientAnchor();
		
		int cellnum = 0;
		Cell cell = row.createCell(cellnum++);

		anchor.setCol1(cell.getColumnIndex());
		anchor.setCol2(cell.getColumnIndex() + 1);
		anchor.setRow1(row.getRowNum());
		anchor.setRow2(row.getRowNum() + 3);
		
		Comment comment = drawing.createCellComment(anchor);
		RichTextString str = factory.createRichTextString("Olá clã");
		comment.setString(str);
		comment.setAuthor("El Picón");
		
		cell.setCellComment(comment);
		
		
		List<Aluno> listaAlunos = new ArrayList<Aluno>();
		listaAlunos.add(new Aluno("Felipe", 7, 8, 0, false));
		listaAlunos.add(new Aluno("André", 5, 8, 0, false));
		listaAlunos.add(new Aluno("Jeniffer", 7, 6, 0, false));
		listaAlunos.add(new Aluno("Cation", 10, 10, 0, false));
		listaAlunos.add(new Aluno("Zeca", 3, 5, 0, false));
		CellStyle style = workbook.createCellStyle();
		style.setFillBackgroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
		style.setFillPattern(CellStyle.BIG_SPOTS);
		style.setBorderBottom(HSSFCellStyle.BORDER_MEDIUM);
		style.setBorderTop(HSSFCellStyle.BORDER_MEDIUM);
		style.setBorderRight(HSSFCellStyle.BORDER_MEDIUM);
		style.setBorderLeft(HSSFCellStyle.BORDER_MEDIUM);
		style.setAlignment(style.ALIGN_CENTER);
		sheetAlunos.addMergedRegion(new CellRangeAddress(1, 1, 1, 2));


		int rownum = 0;
		for (Aluno aluno : listaAlunos) {

			row = sheetAlunos.createRow(rownum++);
			cellnum = 0;
			Cell cellNome = row.createCell(cellnum++);
			cellNome.setCellValue(aluno.getNome());
			Cell cellNota1 = row.createCell(cellnum++);
			cellNota1.setCellValue(aluno.getNota1());
			Cell cellNota2 = row.createCell(cellnum++);
			cellNota2.setCellValue(aluno.getNota2());
			Cell cellMedia = row.createCell(cellnum++);
			cellMedia.setCellValue((aluno.getNota1() + aluno.getNota2()) / 2);
			Cell cellAprovado = row.createCell(cellnum++);
			cellAprovado.setCellValue(cellMedia.getNumericCellValue() >= 6);
			row.getCell(0).setCellStyle(style);
			row.getCell(1).setCellStyle(style);
			row.getCell(2).setCellStyle(style);
			row.getCell(3).setCellStyle(style);
			row.getCell(4).setCellStyle(style);

			try {
				FileOutputStream out = new FileOutputStream(new File(CriarExcel.fileName));
				workbook.write(out);
				out.close();
				System.out.println("Arquivo excel criado com sucesso!");

			} catch (FileNotFoundException e) {
				e.printStackTrace();
				System.out.println("Arquivo não encontrado!");
			} catch (IOException e) {
				e.printStackTrace();
				System.out.println("Erro na edição do arquivo!");
			}
		}
	}
}