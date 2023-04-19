import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;

public class Teste2 {
	
	public static void lerExcel(String caminho_arquivo) throws Exception {
		
		Workbook workbook = new Workbook(caminho_arquivo);
		
		// Obter todas as planilhas
		WorksheetCollection collection = workbook.getWorksheets();
		
		for (int worksheetIndex = 0; worksheetIndex < collection.getCount(); worksheetIndex++) {

			  // Obter planilha usando seu índice
			  Worksheet worksheet = collection.get(worksheetIndex);

			  // Imprimir nome da planilha
			  System.out.print("Worksheet: " + worksheet.getName());

			  // Obter número de linhas e colunas
			  int rows = worksheet.getCells().getMaxDataRow();
			  int cols = worksheet.getCells().getMaxDataColumn();

			  // Percorrer as linhas
			  for (int i = 0; i < rows; i++) {

			    // Percorrer cada coluna na linha selecionada
			    for (int j = 0; j < cols; j++) {
			        //Valor da célula
			    	System.out.print(worksheet.getCells().get(i, j).getValue() + " | ");
			    }
			    // Imprimir quebra de linha
			    System.out.println(" ");
			  }
		}		
	}
	
	public static void escreveEmExcel() throws Exception {
		
		Workbook wkb = new Workbook();

		// Acessa a primeira planilha da pasta de trabalho.
		Worksheet worksheet = wkb.getWorksheets().get(0);

		// Adiciona conteúdo na célula
		
		worksheet.getCells().get("A1").putValue("ColumnA");
		worksheet.getCells().get("B1").putValue("ColumnB");
		worksheet.getCells().get("C1").putValue("ColumnB");
		
		for(int i = 2; i < 5; i++) {
			worksheet.getCells().get("A"+i).putValue("ColumnA"+i);
			worksheet.getCells().get("B"+i).putValue("ColumnB"+i);
			worksheet.getCells().get("C"+i).putValue("ColumnC"+i);
		}
		
		// Salva a pasta de trabalho como arquivo XLSX
		wkb.save("Excel.xlsx");
	}
	
	public static void main(String[] args) throws Exception {
		
		//lerExcel("POP2022_Municipios.xls");
		escreveEmExcel();
	}

}
