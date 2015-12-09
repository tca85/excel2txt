package br.com.ache.lims;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.Iterator;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;


/**
 * Converter arquivo excel em texto (.txt) para ser utilizado via comando na SM69 do SAP
 *  
 * obs: deve ser compilado na versão 1.4.2 do Java
 * 
 * 09.12.2015
 * 
 * @author Thiago Cordeiro Alves / Uderson Fermino
 *
 */
public class ReadExcel {
	public static void main(String[] args) {
//		File dirOrigem = new File("/opt/conversor_lims/origem/");
//		File dirDestino = new File("/opt/conversor_lims/destino/");
		
		File dirOrigem = new File("c:/temp/excel/origem/");
		File dirDestino = new File("c:/temp/excel/destino/");
		
		String separator = System.getProperty("file.separator");

		try {
			Iterator listaArquivosXLS = FileUtils.iterateFiles(dirOrigem, null, false);
			
			if (!listaArquivosXLS.hasNext()) {
				System.out.println("Não existem arquivos XLS para serem convertidos em TXT");
			}			

			while (listaArquivosXLS.hasNext()) {
				File arquivo = ((File) listaArquivosXLS.next());

				if (arquivo.getName().toUpperCase().endsWith(".XLS")) {
					String value = "";

					System.out.println("Lendo aquivo XLS...");

					FileInputStream inputStream = new FileInputStream(arquivo);
					Workbook workbook = new HSSFWorkbook(inputStream);

					Sheet firstSheet = workbook.getSheetAt(2);

					for (int i = 0; i <= firstSheet.getLastRowNum(); i++) {
						Row nextRow = firstSheet.getRow(i);

						if (nextRow.getRowNum() >= 10) {
							for (int j = 0; j <= nextRow.getLastCellNum(); j++) {								
								Cell cell = nextRow.getCell(j);

								if (cell != null && cell.getCellType() != Cell.CELL_TYPE_BLANK) {
									switch (cell.getCellType()) {
									case Cell.CELL_TYPE_STRING:
										value = value + cell.getStringCellValue() + ";";
										break;
									case Cell.CELL_TYPE_BOOLEAN:
										value = value + String.valueOf(cell.getBooleanCellValue()) + ";";
										break;
									case Cell.CELL_TYPE_NUMERIC:
										value = value + String.valueOf(cell.getNumericCellValue()) + ";";
										break;
									}
								}
							}
							value = value + "\n";
						}
					}

					File file = new File(dirDestino.getAbsolutePath() + separator + arquivo.getName().replaceAll(".xls", ".txt"));
					FileWriter fileWriter = null;

					try {
						fileWriter = new FileWriter(file);
						fileWriter.write(value.replaceAll("\n", System.getProperty("line.separator")));
						fileWriter.close();
						System.out.println("Arquivo convertido com sucesso");
					} catch (IOException ex) {
						System.out.println(ex.getMessage());
					} finally {
						try {
							fileWriter.close();
						} catch (IOException ex) {
							System.out.println(ex.getMessage());
						}
					}
					arquivo.delete();
					inputStream.close();
				}
			}

		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
	}
}