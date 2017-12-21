package br.com.assertsistemas.apachepoi.examples;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;

public class AbrirExcel {

	private static final String nomeArquivo = "C:/Users/Marcos/Desktop/teste.xls";

	public static void main(String[] args) {
		List<Aluno> listaAlunos = new ArrayList<Aluno>();

		HSSFWorkbook workbook = null;
		try {
			FileInputStream arquivo = new FileInputStream(new File(AbrirExcel.nomeArquivo));

			try {
				workbook = new HSSFWorkbook(arquivo);
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}


		HSSFSheet sheetAlunos = workbook.getSheetAt(0);
		Iterator<Row> rowIterator = sheetAlunos.iterator();

		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();
			Iterator<Cell> cellIterator = row.cellIterator();

			Aluno aluno = new Aluno();
			listaAlunos.add(aluno);
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				switch (cell.getColumnIndex()) {
				case 0:
					aluno.setNome(cell.getStringCellValue());
					break;
				case 1:
					aluno.setNota1(cell.getNumericCellValue());
					break;
				case 2:
					aluno.setNota2(cell.getNumericCellValue());
					break;
				case 3:
					aluno.setMedia(cell.getNumericCellValue());
					break;
				case 4:
					aluno.setAprovado(cell.getBooleanCellValue());
					break;
				}
			}
		}

		try {
			arquivo.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}catch(

	FileNotFoundException e)
	{
			e.printStackTrace();
			System.out.println("Arquivo não encontrado!");
		}if(listaAlunos.size()==0)
	{
		System.out.println("Nenhum aluno encontrado!");
	}else
	{
		double soma = 0;
		double maior = 0;
		double menor = listaAlunos.get(0).getMedia();
		int aprovados = 0;
		int reprovados = 0;
		for (Aluno aluno : listaAlunos) {
			System.out.println("Aluno: " + aluno.getNome() + "||" + "Média: " + aluno.getMedia());
			soma = soma + aluno.getMedia();
			if (aluno.getMedia() > maior) {
				maior = aluno.getMedia();
			}
			if (aluno.getMedia() < menor) {
				menor = aluno.getMedia();
			}
			if (aluno.getMedia() >= 6) {
				aprovados++;
			}
			if (aluno.getMedia() < 6) {
				reprovados++;
			}
		}
		double media = soma / listaAlunos.size();
		System.out.println("A meédia das notas é: " + media);
		System.out.println("A maior nota é: " + maior);
		System.out.println("A menor nota é: " + menor);
		System.out.println("O número de alunos aprovados é: " + aprovados);
		System.out.println("O número de reprovados é: " + reprovados);
		System.out.println("Número total de alunos: " + listaAlunos.size());
	}
}
}
