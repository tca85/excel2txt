package br.com.ache.lims;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Properties;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * A dirty simple program that reads an Excel file.
 * 
 * @author www.codejava.net
 *
 */
public class ReadExcel {
	public static void main(String[] args) {
		new ReadExcel().buscarArquivos();
	}

	/**
	 * busca arquivos...
	 */
	private void buscarArquivos() {
		List<Properties> listaProperties = new ArrayList<Properties>();

		try {
			buscarArquivoProperties(listaProperties, "resources/excel2txt.properties");

			converterExcelTXT(listaProperties);

		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	/**
	 * Converte arquivo Excel para TXT
	 */
	private void converterExcelTXT(List<Properties> listaProperties) {
		File dirOrigem;
		File dirDestino;
		FileInputStream inputStream;
		
		String separator = System.getProperty("file.separator");

		for (Properties prop : listaProperties) {
			try {
				System.out.println("Verificando se existem arquivos para serem convertidos...");

				dirOrigem = new File(prop.getProperty("dir_origem"));
				dirDestino = new File(prop.getProperty("dir_destino"));

				for (File arquivo : dirOrigem.listFiles()) {
					if (arquivo.getName().toUpperCase().endsWith(".XLS")) {
						System.out.println("Lendo aquivo excel");

						inputStream = new FileInputStream(arquivo);
						Workbook workbook = new HSSFWorkbook(inputStream);

						Sheet firstSheet = workbook.getSheetAt(Integer.parseInt(prop.getProperty("sheet")));
						Iterator<Row> iterator = firstSheet.iterator();

						String value = "";

						while (iterator.hasNext()) {
							Row nextRow = iterator.next();

							if (nextRow.getRowNum() > Integer.parseInt(prop.getProperty("sheet_row"))) {
								Iterator<Cell> cellIterator = nextRow.cellIterator();
								System.out.println("Lendo nova excel do excel");

								while (cellIterator.hasNext()) {
									Cell cell = cellIterator.next();
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

						System.out.println("Gravando aquivo convertido");

						File file = new File(dirDestino.getAbsolutePath() + separator + arquivo.getName().replaceAll(".xls", ".txt"));
						FileWriter fileWriter = null;

						try {
							fileWriter = new FileWriter(file);
							fileWriter.write(value.replaceAll("\n", System.getProperty("line.separator")));
							fileWriter.close();
							System.out.println("Gravando aquivo convertido em txt");
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
						workbook.close();
						inputStream.close();
					}
				}
			} catch (Exception e) {
				System.out.println(e);
				e.printStackTrace();
			}
		}
	}

	/**
	 * 
	 * @param listaProperties
	 * @param nomeArquivo
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	private void buscarArquivoProperties(List<Properties> listaProperties, String nomeArquivo)
			throws IOException, FileNotFoundException {
		Properties propertiesFile = new Properties();
		InputStream ipt1 = getClass().getClassLoader().getResourceAsStream(nomeArquivo);

		if (ipt1 != null) {
			propertiesFile.load(ipt1);
			listaProperties.add(propertiesFile);
			ipt1.close();
		} else {
			throw new FileNotFoundException("arquivo de propriedades " + nomeArquivo + " não encontrado");
		}
	}
}
