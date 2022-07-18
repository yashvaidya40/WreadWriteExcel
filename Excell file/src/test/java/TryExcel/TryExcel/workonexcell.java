package TryExcel.TryExcel;
import java.io.FileInputStream;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class workonexcell {
	static int totalRows;
	static int totalCols;
//	static int i;
//	static int j;
public static String[][] getData(String filepath, String SheetName,int jy) throws Exception {
		
		
		FileInputStream ExcelFile = new FileInputStream(filepath);
		XSSFWorkbook wb = new XSSFWorkbook(ExcelFile);
		XSSFSheet sh = wb.getSheet(SheetName);
		  
		totalCols = sh.getRow(1).getPhysicalNumberOfCells();
		
		String[][] results = new String[1][totalCols];
		
		for (int j = 0; j <= (totalCols-1); j++) {
		
			
			// get table data values
			results[0][j] =String.valueOf(sh.getRow(jy).getCell(j));
			
			 
		}
		
		wb.close();

		
		
		return results;
		
	}
public static void main(String[] args) throws Exception {

	String[][] name = null ;
	int i1=1;
	int j2=3;
	
while(i1<j2)
{
	name = workonexcell.getData("C:\\Users\\Yash\\workspace\\Excel\\New Microsoft Excel Worksheet.xlsx", "Sheet1",i1);
	for (int i=0;i<totalCols;i++)
	{
		System.out.print(name[0][i]);
		if(name[0][i]=="null")
		{
			System.out.println("Enter Data in Excel in "+i1+"Row and "+i+"Coloum");
			
		}
	}
	i1++;
}
}
 
 
 
 

}