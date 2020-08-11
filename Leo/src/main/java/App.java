import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;

import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Connection;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;


public class App {
       private static final String URL = "http://receita.economia.gov.br/orientacao/tributaria/pagamentos-e-parcelamentos/taxa-de-juros-selic#Selicmensalmente";
       private static final String NOME_ARQUIVO_GERADO = "RelatorioLeo.xlsx";
       private static final String NOME_PLANILHA_GERADA = "RelatorioLeo";
       private static final int COLUNAS_TABELA = 9;
       private static final int LINHAS_TABELA = 13;
       private static int contadorRegistros = 0;
       
       
	public static void main(String[] args) {
		
		String[][] tabelaStringFormatada = new String[3][500];
		Document htmlPagina = getDocumento(URL);		
		
		//busca a 3a tabela da página 1995-2003
		Element tabelaHTML = htmlPagina.select("table").get(3); 
		getTabela(tabelaHTML,tabelaStringFormatada);
		
		//adiciona a 2a tabela 2004-2012
		tabelaHTML = htmlPagina.select("table").get(2); 
		getTabela(tabelaHTML,tabelaStringFormatada);
		
		//adiciona a 1a tabela 2013-2020
		tabelaHTML = htmlPagina.select("table").get(1); 
		getTabela(tabelaHTML,tabelaStringFormatada);
		
		imprimeExcel(tabelaStringFormatada);
       }
	
	// Busca documento HTML na URL
	public static Document getDocumento(String URL) {
		try {

			Connection.Response loginForm = Jsoup.connect(URL)
							            .method(Connection.Method.GET)
							            .execute();
			
			return loginForm.parse();
			
		} catch (IOException e) {
			e.printStackTrace();
		}
		return null;
       }
	
	// Converte os dados html para array de string formatado
	public static void getTabela(Element tabelaHTML, String[][] tabelaFormatada) {
		
		Elements rows = tabelaHTML.select("tr");
		String tabelaDesformatada[][] = new String[LINHAS_TABELA][COLUNAS_TABELA];
		
		// Monta tabela simples/desformatada com layout igual ao table
		for (int i = 0; i < rows.size(); i++) {
		    Element row = rows.get(i);
		    Elements cols = row.select("td");
		    
		    for (int y = 0; y < 9; y++) {
		    	tabelaDesformatada[i][y] = cols.get(y).text();
		    }
		}
		
		// Monta tabela com layout final
		for (int i = 1; i < COLUNAS_TABELA; i++) {  // colunas
			for (int y = 1; y < LINHAS_TABELA; y++) { //linhas
				
				if (tabelaDesformatada[y][i].toString().trim().isEmpty() == true)
				{
					break;
				}

				tabelaFormatada[0][contadorRegistros] = tabelaDesformatada[0][i].toString();
				tabelaFormatada[1][contadorRegistros] = getNumeroMes(tabelaDesformatada[y][0].toString());
				tabelaFormatada[2][contadorRegistros] = tabelaDesformatada[y][i].toString();
				
				contadorRegistros++;
			}
		}
	}
	
	// Printa Array de String final e completo no Excel
	public static boolean imprimeExcel(String[][] tabelaString) {
		
		// Abre o arquivo em memória
		Workbook wb = new XSSFWorkbook();
		CreationHelper createHelper = wb.getCreationHelper();
		org.apache.poi.ss.usermodel.Sheet sheet = wb.createSheet(NOME_PLANILHA_GERADA);
		
		for (int i = 0; i < contadorRegistros; i++) {
			
			// Cria linha
			Row row = sheet.createRow(i);
			
			// Preenche cada célula, 3 colunas
			row.createCell(0).setCellValue(createHelper.createRichTextString(tabelaString[0][i]).toString());
			row.createCell(1).setCellValue(createHelper.createRichTextString(tabelaString[1][i]).toString());
			row.createCell(2).setCellValue(createHelper.createRichTextString(tabelaString[2][i]).toString());
		}

		// Escreve o arquivo
		try (OutputStream fileOut = new FileOutputStream(NOME_ARQUIVO_GERADO)) {
			wb.write(fileOut);
			wb.close();
		} catch (IOException e) {
			e.printStackTrace();
		}

		return true;
	}
	
	// Utilitário: Retorna o número do mês a partir do Nome com localização
	public static String getNumeroMes (String nomeMes) {
		
		SimpleDateFormat inputFormat = new SimpleDateFormat("MMMM");
		
		Calendar cal = Calendar.getInstance();
		
		try {
			cal.setTime(inputFormat.parse(nomeMes));
		} catch (ParseException e) {
			e.printStackTrace();
		}
		
		SimpleDateFormat outputFormat = new SimpleDateFormat("MM");
		
		return outputFormat.format(cal.getTime());
	}
}