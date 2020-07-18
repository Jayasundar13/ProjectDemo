package testing.demo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class Sample {
	public static void main(String[] args) throws IOException {
		
		File excelLoc = new File("C:\\Users\\MUGUNTH\\eclipse-workspace\\testing.demo\\ExcelSheet\\Book1.xlsx");
		
		FileInputStream stream = new FileInputStream(excelLoc);
		
		Workbook w = new XSSFWorkbook(stream);
		
		Sheet s = w.getSheet("Sheet1");
		
		for (int i = 0; i < s.getPhysicalNumberOfRows(); i++) {
		
		Row r = s.getRow(i);
		
		for (int j = 0; j < r.getPhysicalNumberOfCells(); j++) {
		
		Cell c = r.getCell(j);
		
		int type = c.getCellType();
		if (type == 1) {
			
			String name = c.getStringCellValue();
			System.out.println(name);
			
		}if (type == 0) {
			
			boolean b = DateUtil.isCellDateFormatted(c);
			if(b==true) {
				Date date = c.getDateCellValue();
				SimpleDateFormat fr =new SimpleDateFormat("dd-MMM-yy");
				String dob = fr.format(date);
				System.out.println(dob);
				
			}else {
			
			double d = c.getNumericCellValue();
			long l = (long)d;
			
			String data = String.valueOf(l);
			System.out.println(data);
			
		}
		
		}
		
		}
	}
}
	
}

