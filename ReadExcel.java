package selTest;

// this package will help to read specific records in a excel
import java.io.File;
import java.io.FileInputStream;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {

	
	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub

		//create a object of file class and specify pathof excel sheet, import related class "java.io"
		
		File src = new File("C:\\Users\\suparna_das01\\workspace\\DemoProject\\TestSuite\\testdata\\Testdata.xlsx");
		
		//1.create a object of class fileinputstream and pass the file source to it, import related class from "java.io"
		//2. suggests to throw exception, use normal exception
		
		FileInputStream fis = new FileInputStream(src);
		
		//create a workbook object 'Wb' and passon the fileinputstream value to it.
		
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		
		//create a worksheet object 'Sheet1' and specify which sheet to look for, here it is sheet1 which starts at index value 0
		
		XSSFSheet sheet1= wb.getSheetAt(0);
		
		//Fetch the string value of row1 and column 0 from sheet1 and assign it a string variable type
		String data0=sheet1.getRow(1).getCell(0).getStringCellValue();
		
		//print the output
		System.out.println("Data from Excel is :"+ data0);
		
		//Fetching multiple values
		String data1=sheet1.getRow(1).getCell(1).getStringCellValue();
		
		System.out.println("Data from Excel is :"+ data1);
		
		//Close the workbook to take care of memoryleak warning message
		wb.close();
		
		
	}

}
