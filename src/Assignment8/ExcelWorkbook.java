package Assignment8;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.File;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWorkbook {

	public static void main(String[] args) {
//		XSSFWorkbook workbook = new XSSFWorkbook();
//		XSSFSheet sheet = workbook.createSheet("sheet1");
//		Row row = sheet.createRow(0);
//		Cell cell = row.createCell(0);
//		try {
//		    FileOutputStream out = new FileOutputStream(
//		            new File("Test.xlsx"));
//		    workbook.write(out);
//		    out.close();
//		} catch (Exception e) {
//		    e.printStackTrace();
//		}
		
		try(XSSFWorkbook workbook=new XSSFWorkbook()) {
			
			XSSFSheet sheet = workbook.createSheet("sheet3");
//			Sheet sheet=workbook.createSheet("Sheet1");
						
			Object[][] data= {
					{"Name","Age","Email"},
					{"John Doe",30,"john@test.com"},
					{"Jain Doe",28,"jane@test.com"},
					{"Bob Smith",35,"bob@example.com"},
					{"Swapnil",37,"swapnil@example.com"},
						
			};
			
			int rowNum=0;
			for(Object[] rowdata:data) {				
				Row row=sheet.createRow(rowNum++);				
				int colNum=0;
				for(Object field:rowdata) {
					Cell cell=row.createCell(colNum++);
					if(field instanceof String) {
						cell.setCellValue((String)field);
					}else if(field instanceof Integer) {
						cell.setCellValue((Integer)field);
					}
				}				
			}
			try{
				FileOutputStream os=new FileOutputStream("Test3.xlsx");
				workbook.write(os);
				os.close();
			}catch (Exception e) {
				e.printStackTrace();
			}			
			System.out.println("Data Added Successfully to file...");
			workbook.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
		

	}

}
