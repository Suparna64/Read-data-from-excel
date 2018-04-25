package selTest;

//this package will help to read all the rows and columns in a excel using Iterator

import java.io.File;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import java.io.IOException;
import java.util.Iterator;

public class ReadExcelIterator {
	
	public static final String FILE_PATH = "C:\\Users\\suparna_das01\\workspace\\DemoProject\\TestSuite\\testdata\\Testdata.xlsx";
	
	public static void main(String[] args) throws IOException, Throwable, InvalidFormatException {
		
		// Creating a Workbook from an Excel file (.xls or .xlsx)
		
			Workbook workbook = WorkbookFactory.create(new File(FILE_PATH));
		
			// Retrieving the number of sheets in the Workbook
	        System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");

	        /* Iterating over all the sheets in the workbook using sheetIterator and iterate over it */

	        Iterator<Sheet> sheetIterator = workbook.sheetIterator();
	        System.out.println("Retrieving Sheets using Iterator");
	        while (sheetIterator.hasNext()) {
	            Sheet sheet = sheetIterator.next();
	            System.out.println("=> " + sheet.getSheetName());
	        }
		
		      /*Iterating over all the rows and columns in a Sheet */

	             // Getting the Sheet at index zero
	             Sheet sheet = workbook.getSheetAt(0);

	             // Create a DataFormatter to format and get each cell's value as String
	             DataFormatter dataFormatter = new DataFormatter();

	             // Use the rowIterator and columnIterator and iterate over them
	             System.out.println("\n\nIterating over Rows and Columns using Iterator\n");
	             Iterator<Row> rowIterator = sheet.rowIterator();
	             while (rowIterator.hasNext()) {
	                 Row row = rowIterator.next();

	                 // Now let's iterate over the columns of the current row
	                 Iterator<Cell> cellIterator = row.cellIterator();

	                 while (cellIterator.hasNext()) {
	                     Cell cell = cellIterator.next();
	                     String cellValue = dataFormatter.formatCellValue(cell);
	                     System.out.print(cellValue + "\t");
	                 }
	                 System.out.println();
	             }
	        
		//Close the workbook to take care of memoryleak warning message
	             workbook.close();
		
		
	}

}
