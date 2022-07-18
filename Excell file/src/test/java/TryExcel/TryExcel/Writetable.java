package TryExcel.TryExcel;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import io.github.bonigarcia.wdm.WebDriverManager;

//private XSSFSheet sheet = null;
//int index = workbook.getSheetIndex(sheetName);
//sheet = workbook.getSheetAt(index);
//C:\Users\Yash\Downloads\chromedriver_win32 (2)
public class Writetable {

	public static void main(String[] args) throws IOException {
		WebDriver driver;
		WebDriverManager.chromedriver().setup();
		driver = new ChromeDriver();   
		//System.setProperty("webdriver.chrome.driver", System.getProperty("user.dir") + "\\driver\\chromedriver.exe");
		//WebDriver driver = new ChromeDriver();
		driver.get("https://www.w3schools.com/html/html_tables.asp");
		String beforxpath = "//*[@id='customers']//tr[";
		String[] str = {"yash", "CompanyName","CompanyContact","Country"};
		
		List<WebElement> rows = driver.findElements(By.xpath("//*[@id='customers']//tr"));
		List<WebElement> col = driver.findElements(By.xpath("//*[@id='customers']//th"));
		String path = System.getProperty("user.dir") + "\\ExcelTestdata\\Write.xlsx";
		
		Xls_Reader reader = null ;
		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet sh = wb.createSheet("sheet");
		XSSFRow row = sh.createRow(0);
		//XSSFCell cell= row.createCell(2);
		File tempFile = new File(path);

		if(!tempFile.exists()){ 
		FileOutputStream fout = new FileOutputStream(path);
		wb.write(fout);
		reader = new Xls_Reader(path);
		}else	
		 reader = new Xls_Reader(path);
		
		if (reader.isSheetExist("Tabledata"))
			reader.removeSheet("Tabledata");
			
		
		reader.addSheet("Tabledata");
		//reader.addColumn("Tabledata", "CompanyName");
		reader.addColumn("Tabledata", "CompanyContact");
		//reader.addColumn("Tabledata", "Country");
		String companyname = null;
		
		for (int j = 1; j <=col.size() ; j++) {
			for (int i = 2; i <= rows.size(); i++) {
				String str1=str[j];
				
				String afterXpath = "]/td["+j+"]";
				String actual = beforxpath + i + afterXpath;
				
				companyname = driver.findElement(By.xpath(actual)).getText();
				reader.setCellData("Tabledata", str1, i, companyname);
				
			}
		}
		System.out.println("End");
		System.out.println(reader.getCellData("Tabledata", 2, 2));
		driver.close();
	}

}
