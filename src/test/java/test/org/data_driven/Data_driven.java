package test.org.data_driven;

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

public class Data_driven {

	public static void main(String[] args) throws IOException {
		File f = new File("C:\\Users\\Ashwin Babu\\Documents\\Selenium Learning\\Demo.xlsx");
		FileInputStream f1 = new FileInputStream(f);
		Workbook w= new XSSFWorkbook(f1);
		Sheet sheet = w.getSheet("Sheet");
		for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
			Row row = sheet.getRow(i);
			for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
				Cell cell = row.getCell(j);
				int cellType = cell.getCellType();
				String value = "";
				if (cellType ==1) {
					value = cell.getStringCellValue();
					}
				else if (DateUtil.isCellDateFormatted(cell)) {
					 Date dateCellValue = cell.getDateCellValue();
					 SimpleDateFormat sim = new SimpleDateFormat("MM/dd/yyyy");
					 value = sim.format(dateCellValue);
					
				}
				
				else {
					double numericCellValue = cell.getNumericCellValue();
					long l = (long) numericCellValue;
					value = String.valueOf(l);
				}
				
				
			}
			
		}
		
		

	}

}
