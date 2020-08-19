package relatorio.leo;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeFormatterBuilder;
import java.time.temporal.ChronoField;
import java.util.Arrays;
import java.util.HashSet;
import java.util.Locale;
import java.util.Optional;
import java.util.Set;
import java.util.TreeSet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Connection;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class App {

  private static final class Registro implements Comparable<Registro> {
    int ano;
    int mes;
    double porcentagem;

    public Registro(int ano, int mes, double porcentagem) {
      this.ano = ano;
      this.mes = mes;
      this.porcentagem = porcentagem;
    }

    @Override
    public int compareTo(Registro o) {
      int compare = Integer.compare(ano, o.ano);
      if (compare == 0) {
        compare = Integer.compare(mes, o.mes);
      }
      return compare;
    }
  }

  private static final Logger logger = LoggerFactory.getLogger(App.class);
  private static final Locale PT_BR = Arrays.stream(Locale.getAvailableLocales())
      .filter(locale -> locale.toString().equals("pt_BR"))
      .findFirst()
      .orElseThrow();
  private static final DateTimeFormatter FORMATTER_MES = new DateTimeFormatterBuilder().parseCaseInsensitive()
      .appendPattern("MMMM")
      .toFormatter(PT_BR);

  private static final String URL = "http://receita.economia.gov.br/orientacao/tributaria/pagamentos-e-parcelamentos/taxa-de-juros-selic#Selicmensalmente";
  private static final String NOME_ARQUIVO_GERADO = "RelatorioLeo2.xlsx";
  private static final String NOME_PLANILHA_GERADA = "RelatorioLeo";
  private static final int COLUNAS_TABELA = 9;
  private static final int LINHAS_TABELA = 13;

  public static void main(String[] args) {
    Document htmlPagina = getDocumento(URL).orElseThrow();
    Set<Registro> todosRegistros = new TreeSet<>();

    for (int a1 = 1; a1 <= 3; a1++) {
      todosRegistros.addAll(getTabela(htmlPagina.select("table").get(a1)));
    }

    imprimeExcel(todosRegistros);
  }

  // Busca documento HTML na URL
  public static Optional<Document> getDocumento(String url) {
    try {
      return Optional.ofNullable(Jsoup.connect(URL).method(Connection.Method.GET).execute().parse());
    } catch (IOException e) {
      logger.error("Pegando URL {}", url, e);
    }
    return Optional.empty();
  }

  /**
   * Converte a tabela em HTML para os registros interessantes para o programa
   * 
   * @param tabelaHTML
   * @param tabelaFormatada
   */
  public static Set<Registro> getTabela(Element tabelaHTML) {
    Set<Registro> resposta = new HashSet<>();

    Elements rows = tabelaHTML.select("tr");

    String[][] tabelaDesformatada = new String[COLUNAS_TABELA][LINHAS_TABELA];

    // Monta tabela simples/desformatada com layout igual ao table
    for (int i = 0; i < rows.size(); i++) {
      Element row = rows.get(i);
      Elements cols = row.select("td");

      for (int y = 0; y < 9; y++) {
        tabelaDesformatada[i][y] = cols.get(y).text();
      }
    }

    // Monta tabela com layout final
    for (int i = 1; i < COLUNAS_TABELA; i++) { // colunas
      for (int y = 1; y < LINHAS_TABELA; y++) { // linhas

        if (tabelaDesformatada[y][i].toString().trim().isEmpty() == true) {
          break;
        }

        resposta.add(new Registro(Integer.parseInt(tabelaDesformatada[0][i].toString()),
            getNumeroMes(tabelaDesformatada[y][0].toString()),
            Double.parseDouble(tabelaDesformatada[y][i].toString().replace(',', '.'))));

      }
    }

    return resposta;
  }

  // Printa Array de String final e completo no Excel
  public static boolean imprimeExcel(Set<Registro> registros) {

    // Abre o arquivo em memória
    Workbook wb = new XSSFWorkbook();

    org.apache.poi.ss.usermodel.Sheet sheet = wb.createSheet(NOME_PLANILHA_GERADA);

    int rowNum = 0, colNum = 0;
    Row row = sheet.createRow(rowNum++);
    row.createCell(colNum++).setCellValue("Ano");
    row.createCell(colNum++).setCellValue("Mês");
    row.createCell(colNum++).setCellValue("Taxa (%)");

    for (Registro registro : registros) {
      row = sheet.createRow(rowNum++);
      colNum = 0;

      row.createCell(colNum++).setCellValue(registro.ano);
      row.createCell(colNum++).setCellValue(registro.mes);
      row.createCell(colNum++).setCellValue(registro.porcentagem);
    }

    // Escreve o arquivo
    try (OutputStream fileOut = new FileOutputStream(NOME_ARQUIVO_GERADO)) {
      wb.write(fileOut);
      return true;
    } catch (IOException e) {
      logger.error("Escrevendo arquivo", e);
    }
    return false;

  }

  /**
   * Utilitário: Retorna o número do mês a partir do Nome com localização
   * 
   * @param nomeMes
   * @return
   */
  public static int getNumeroMes(String nomeMes) {
    return FORMATTER_MES.parse(nomeMes).get(ChronoField.MONTH_OF_YEAR);
  }
}