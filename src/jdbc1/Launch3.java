package jdbc1;
import java.io.File;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
public class Launch3 {
	
	public static void main(String[] args) {
		 
	
			String url="jdbc:oracle:thin:@localhost:1521:XE";
			String un="system";
			String pw="system";
			Connection con=null;
			Statement stmt=null;
			ResultSet res=null;
			//Create blank workbook
		      HSSFWorkbook workbook = new HSSFWorkbook();
		      
		      //Create a blank sheet
		      HSSFSheet sheet = workbook.createSheet( "student4");

		    
		      		      
		      
			try
			{
				Class.forName("oracle.jdbc.driver.OracleDriver");
				System.out.println("Driver loaded successfully");
			}
			catch(Exception e)
			{
				System.out.println("driver problrm");
			}
		try
			{
				con=DriverManager.getConnection(url,un,pw);
				System.out.println("connection established");
			}
			catch(Exception e)
			{
				e.printStackTrace();
			}
	try
	{
		stmt=con.createStatement();
		res=stmt.executeQuery("select * from Student ORDER BY rank ASC");
		
		
		System.out.println("Query Execute Success");
	}
	catch(Exception e)
	{
		e.printStackTrace();
		
	}
	try
	{
		  HSSFRow row=sheet.createRow(0);
	      HSSFCell cell;
	      cell=row.createCell(0);
	      cell.setCellValue("ID");
	      cell=row.createCell(1);
	      cell.setCellValue("NAME");
	      cell=row.createCell(2);
	      cell.setCellValue("MARKS1");
	      cell=row.createCell(3);
	      cell.setCellValue("MARKS2");
	      cell=row.createCell(4);
	      cell.setCellValue("AVG");
	      cell=row.createCell(5);
	      cell.setCellValue("RANK");
	      
		 int i=1;
		 
		while(res.next()==true)
		{
			row=sheet.createRow(i);
			cell=row.createCell(0);
			cell.setCellValue(res.getInt(1));
			cell=row.createCell(1);
			cell.setCellValue(res.getString(2));
			cell=row.createCell(2);
			cell.setCellValue(res.getInt(3));
			cell=row.createCell(3);
			cell.setCellValue(res.getInt(4));
			cell=row.createCell(4);
			cell.setCellValue(res.getInt(5));
			cell=row.createCell(5);
			cell.setCellValue(res.getInt(6));
		
			i++;	
		}
		Row rowtotal=sheet.createRow(8);
		Cell celltotal=rowtotal.createCell(1);
        celltotal.setCellValue("TOTAL:");
      
        Cell cellTotal = rowtotal.createCell(2);
        cellTotal.setCellFormula("SUM(C2:C7)");
         cellTotal = rowtotal.createCell(3);
        cellTotal.setCellFormula("SUM(D2:D7)");
        cellTotal=rowtotal.createCell(4);
        cellTotal.setCellFormula("AVERAGE(E2:E7)");
      
		 
	FileOutputStream out=new FileOutputStream(new File("student5.xls"));
	workbook.write(out);
	out.close();
	System.out.println("student4.xls is successfully written");
	}
	
	
	catch(Exception e)
	{
		e.printStackTrace();
	}
}
}
