package selTest;

//this package will help to read all records in a excel using loop
import java.io.File;
import java.io.FileInputStream;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelAlldata {

	
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
		//Sheet1 is initiated
		XSSFSheet sheet1= wb.getSheetAt(0);
		//get the count of number of rows in the sheet
		int rowcount =sheet1.getLastRowNum();
		
		//print the total rows
		System.out.println("Total number of rows is:"+ rowcount);
		
		//using loop to read all row values
		for (int i=0; i< rowcount; i++)
		{
			
			String data0=sheet1.getRow(i).getCell(0).getStringCellValue();
			
			System.out.println("Data from Row"+ i +"is"+ data0);
			
		}
		//Close the workbook to take care of memoryleak warning message
		wb.close();
		
		
	}

}
