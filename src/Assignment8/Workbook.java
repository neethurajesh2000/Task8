package Assignment8;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.File;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class Workbook {

	public static void main(String[] args) {
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("sheet1");
		Row row = sheet.createRow(0);		Cell cell = row.createCell(0);
		try {
		    FileOutputStream out = new FileOutputStream(
		            new File("Test.xlsx"));
	    workbook.write(out);
	    out.close();
		} catch (Exception e) {
		    e.printStackTrace();
		}
	}

}
