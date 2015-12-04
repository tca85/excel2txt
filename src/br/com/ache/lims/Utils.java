package br.com.ache.lims;

import java.io.FileInputStream;
import java.io.IOException;
import java.net.FileNameMap;
import java.net.URLConnection;
import java.util.Properties;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Utils {
	public static Properties getProp(String nmeArquivo) throws IOException {
		Properties props = new Properties();
		FileInputStream file = new FileInputStream(nmeArquivo);
		props.load(file);
		return props;
	}

	/**
	 * Verifica atraves de expressao regular se existe a String .XML no texto
	 * passado
	 * 
	 * @param valor
	 * @return
	 */
	public static boolean validaContemStringXML(String valor) {

		Pattern pegaJava = Pattern.compile(".XML");
		// primeiro Java esta em maiusculo!
		Matcher m = pegaJava.matcher(valor.toUpperCase());

		// enquanto o Matcher encontrar o pattern na String fornecida:
		while (m.find()) {
			return true;
		}
		return false;
	}

	/**
	 * Verifica atraves de expressao regular se existe a String .ZIP no texto
	 * passado
	 * 
	 * @param valor
	 * @return
	 */
	public static boolean validaContemStringCompactado(String valor) {

		Pattern pegaZip = Pattern.compile(".ZIP");
		Pattern pega7z = Pattern.compile(".7Z");
		Pattern pegaRar = Pattern.compile(".RAR");

		Matcher z = pegaZip.matcher(valor.toUpperCase());
		Matcher s = pega7z.matcher(valor.toUpperCase());
		Matcher r = pegaRar.matcher(valor.toUpperCase());

		// enquanto o Matcher encontrar o pattern na String fornecida:
		while (z.find() || s.find() || r.find()) {
			return true;
		}
		return false;
	}

	public static String getMimeType(String fileUrl) throws java.io.IOException {
		FileNameMap fileNameMap = URLConnection.getFileNameMap();
		String type = fileNameMap.getContentTypeFor(fileUrl);

		return type;
	}
}
