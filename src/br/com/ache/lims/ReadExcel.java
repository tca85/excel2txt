package br.com.ache.lims;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Properties;

import org.apache.log4j.Logger;
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

	private static Logger logger = Logger.getLogger(ReadExcel.class);
	private File dirOrigem = null;
	private File dirDestino = null;
	private FileInputStream inputStream = null;
	private String separator = System.getProperty("file.separator");

	public static void main(String[] args) {
		new ReadExcel();
	}

	public ReadExcel() {

		logger.info("Iniciando o conversor");

		try {
			List<Properties> diretorio = new ArrayList<Properties>();
			File raiz = new File(".");

			System.out.println("Verificando arquivos properties");

			for (File arq : raiz.listFiles()) {
				String nome = arq.getName();
				
				if ((nome.startsWith("excel2txt")) && (nome.endsWith(".properties"))) {
					diretorio.add(Utils.getProp(nome));
					System.out.println("Arquivo de configuração encontrado");
				}
			}

			for (Properties prop : diretorio) {
				try {

					logger.info("Verificando pastas");
					System.out.println("Verificando pastas");
					
					dirOrigem = new File(prop.getProperty("dir_origem"));
					dirDestino = new File(prop.getProperty("dir_destino"));

					for (File arquivo : dirOrigem.listFiles()) {
						if (arquivo.getName().toUpperCase().endsWith(".XLS")) {

							logger.info("Lendo aquivo excel");
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

									logger.info("Lendo nova excel do excel");
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

							logger.info("Gravando aquivo convertido");
							System.out.println("Gravando aquivo convertido");
							
							File file = new File(dirDestino.getAbsolutePath() + separator + arquivo.getName().replaceAll(".xls", ".txt"));
							FileWriter fileWriter = null;

							try {
								fileWriter = new FileWriter(file);
								fileWriter.write(value.replaceAll("\n", System.getProperty("line.separator")));
								fileWriter.close();
								System.out.println("Gravando aquivo convertido em txt");
							} catch (IOException ex) {
								logger.info(ex);
							} finally {
								try {
									fileWriter.close();
								} catch (IOException ex) {
									logger.info(ex);
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
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
