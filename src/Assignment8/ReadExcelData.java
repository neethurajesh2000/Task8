package Assignment8;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class ReadExcelData {

	public static void main(String[] args) {
		
String filepath="Test3.xlsx";
		
		try {
			
			FileInputStream inpstrean=new FileInputStream(filepath);			
			XSSFWorkbook workbook=new XSSFWorkbook(inpstrean);			
//			Sheet sheet=workbook.getSheetAt(0);
			XSSFSheet sheet = workbook.getSheetAt(0);
			for(Row row:sheet) {				
				for(Cell cell:row) {
					
					if(cell.getCellType()==CellType.STRING) {
						System.out.print(cell.getStringCellValue()+"\t\t");
					}else if(cell.getCellType()==CellType.NUMERIC) {
						System.out.print(cell.getNumericCellValue()+"\t\t");
					}
				}
				System.out.println();
			}			
			
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}
