package exelReader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {

	public static void main(String[] args) {
		System.out.println("xlnx parser start");
		
		File xlnxFile = new File("C:\\Users\\Tomek\\OneDrive\\Public\\WorkSpace\\Eclipse\\exelReader\\src\\main\\java\\files\\Zeszyt1.xlsx"); 
		FileInputStream xlnxStream;
		XSSFWorkbook workbook = null;
		try {
			xlnxStream = new FileInputStream(xlnxFile);
			workbook = new XSSFWorkbook(xlnxStream);
		} catch (FileNotFoundException e) {		
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		System.out.println(workbook.getNumberOfSheets());
		XSSFSheet sheet = workbook.getSheetAt(0);
		XSSFSheet sheet2 = workbook.getSheetAt(1);
		
		Iterator<Row> rowIterator = sheet.iterator();
		while(rowIterator.hasNext()) {
			Row row = rowIterator.next();
			Iterator<Cell> cellIterator = row.cellIterator();
			while(cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				
			    if(cell.getCellTypeEnum() == CellType.STRING) {
				    System.out.print(cell.getStringCellValue());
			    }else if(cell.getCellTypeEnum() == CellType.NUMERIC) {
				    System.out.print(cell.getNumericCellValue());
			    }else if(cell.getCellTypeEnum() == CellType.BOOLEAN) {
				    System.out.print(cell.getBooleanCellValue());
			    }						
			}
			System.out.println();
		}

	}
}
