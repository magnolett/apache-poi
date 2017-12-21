package br.com.assertsistemas.apachepoi.practice;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class AbraExcel {

	private static final String fileName = "C:/Users/Marcos/Desktop/Planilhas/a.xls";

	public static void main(String[] args) {
		try(FileInputStream file = new FileInputStream(new File(AbraExcel.fileName));
				HSSFWorkbook workbook = new HSSFWorkbook(file)){
			
			final List<Pessoa> listaPessoas = new ArrayList<>();
			final HSSFSheet sheetPessoas = workbook.getSheetAt(0);
			final Iterator<Row> rows = sheetPessoas.iterator();

			int index = 0;
			while (rows.hasNext()) {
				final Row row = rows.next();
				
				if(index > 2){
					final Pessoa pessoa = new Pessoa();
					row.forEach(cell -> {
						switch (cell.getColumnIndex()) {
						case 0:
							pessoa.setNome(cell.getStringCellValue());
							break;
						case 1:
							pessoa.setIdade((int) cell.getNumericCellValue());
							break;
						case 2:
							pessoa.setCargo(cell.getStringCellValue());
							break;
						case 3:
							pessoa.setSalario(cell.getNumericCellValue());
						}
						
					});
					listaPessoas.add(pessoa);
				}
				
				index++;
			}
		}catch (Exception e) {
			e.printStackTrace();
		}
		main02();
	}
	
	public static void main02() {
		List<Pessoa> listaPessoas = new ArrayList<Pessoa>();

		HSSFWorkbook workbook = null;

		try {
			FileInputStream file = new FileInputStream(new File(AbraExcel.fileName));

			try {
				workbook = new HSSFWorkbook(file);
			} catch (IOException e) {
				e.printStackTrace();
			}

			HSSFSheet sheetPessoas = workbook.getSheetAt(0);
			Iterator<Row> rowIterator = sheetPessoas.iterator();
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				Iterator<Cell> cellIterator = row.cellIterator();

				Pessoa pessoa = new Pessoa();
				listaPessoas.add(pessoa);
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					
					switch (cell.getColumnIndex()) {
					case 0:
						pessoa.setNome(cell.getStringCellValue());
						break;
					case 1:
						pessoa.setIdade((int) cell.getNumericCellValue());
						break;
					case 2:
						pessoa.setCargo(cell.getStringCellValue());
						break;
					case 3:
						pessoa.setSalario(cell.getNumericCellValue());
					}

				}
			}
			try {
				file.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			System.out.println("Arquivo não encontrado!");
		}
		if (listaPessoas.size() == 0) {
			System.out.println("Nenhuma pessoa encontrada!");
		} else {
			double soma = 0;
			double mediaIdade = 0;
			double salario = 0;
			double media = 0;
			for (Pessoa pessoa : listaPessoas) {
				System.out.println("Pessoa: " + pessoa.getNome() + "||" + "Idade: " + pessoa.getIdade() + "||"
						+ "Cargo: " + pessoa.getCargo() + "||" + "Salário: " + pessoa.getSalario());
				soma = soma + pessoa.getIdade();
				mediaIdade = soma / listaPessoas.size();
				salario = salario + pessoa.getSalario();
				media = salario / listaPessoas.size();
			}
			System.out.println("A soma das idades é: " + soma);
			System.out.println("A média das idades é: " + mediaIdade);
			System.out.println("O número total de pessoas é: " + listaPessoas.size());
			System.out.println("A soma do salário dessas pessoas é: " + salario);
			System.out.println("A média salarial dessas pessoas é: " + media);

		}
	}
}