package br.com.sankhya.teste;

	import org.apache.poi.hssf.usermodel.HSSFWorkbook;
	import org.apache.poi.ss.usermodel.Cell;
	import org.apache.poi.ss.usermodel.Row;
	import org.apache.poi.ss.usermodel.Sheet;
	import org.apache.poi.ss.usermodel.Workbook;
	import org.apache.poi.xssf.usermodel.XSSFWorkbook;

	import java.io.FileInputStream;
	import java.io.InputStream;

	import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_NUMERIC;

	public class TesteExcelAula1 {

	    public static void main(String[] args) throws Exception {

	        String diretorioDoArquivo = "C:\\Users\\joao.pedro\\Downloads\\modelo_excel.xls";
	        try{
	            InputStream file = new FileInputStream(diretorioDoArquivo);

	            Workbook workbook;
	            if(diretorioDoArquivo.toLowerCase().endsWith(".xls")){
	                System.out.println("Lendo arquvio XLS");
	                workbook = new HSSFWorkbook(file); // quando for xls
	            } else if (diretorioDoArquivo.toLowerCase().endsWith(".xlsx")){
	                System.out.println("Lendo arquvio XLSX");
	                workbook = new XSSFWorkbook(file); // quando for xlsx
	            }else {
	                throw new Exception("O formato do arquivo não é um arquivo de excel");
	            }

	            Sheet sheet = workbook.getSheetAt(0);
	            // percorrer as linhas
	            for (Row row :sheet){
	                // percorrer as colunas
	                for (Cell cell : row){
	                    switch (cell.getCellType()){
	                        case Cell.CELL_TYPE_NUMERIC:
	                            System.out.println(cell.getNumericCellValue()+ "\t");
	                            break;
	                        case Cell.CELL_TYPE_STRING:
	                            System.out.println(cell.getStringCellValue()+ "\t");
	                            break;
	                        case Cell.CELL_TYPE_BOOLEAN:
	                            System.out.println(cell.getBooleanCellValue()+ "\t");
	                            break;
	                        default:
	                            throw new Exception("Não tratado");
	                    }
	                }
	            }
	            file.close();
	        }catch (Exception e){
	            e.printStackTrace();
	        }

	        System.out.println("Hello");
	    }
}
