package jdbc1;





import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//import org.apache.poi.ss.usermodel.Row;
//import org.apache.poi.xssf.usermodel.XSSFRow;
//import org.apache.poi.xssf.usermodel.XSSFSheet;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class Launch4 {
	
	public static void main(String[] args) {
		
		 
	
			String url="jdbc:oracle:thin:@localhost:1521:XE";
			String un="system";
			String pw="system";
			Connection con=null;
			Statement stmt=null;
			ResultSet res=null;
			//Create blank workbook
		     // XSSFWorkbook workbook = new XSSFWorkbook();
		      
		      //Create a blank sheet
		      //XSSFSheet sheet = workbook.createSheet( "studentdetails ");

		      //Create row object
		      
		     // Row row=sheet.createRow(5);
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
		while(res.next()==true)
		{
			int a=res.getInt(1);
			String b=res.getString(2);
			int c=res.getInt(3);
			int d=res.getInt(4);
			float e=res.getFloat(5);
			int f=res.getInt(6);
//			row.createCell(0).setCellValue(a);
//			row.createCell(1).setCellValue(b);
//			row.createCell(2).setCellValue(c);
//			row.createCell(3).setCellValue(d);
//			row.createCell(4).setCellValue(e);
//			row.createCell(5).setCellValue(f);

			System.out.println("ID"+" "+"NAME"+" "+"MARKS1"+" "+"MARKS2"+" "+"AVG"+" "+"RANK");
			System.out.println(a+"  "+b+"  "+c+"  "+d+"  "+e+"  "+f);
		    	      
			
		}
	}
	catch(Exception e)
	{
		e.printStackTrace();
	}
//		String pr="studentdata.xlsx";
//		FileOutputStream fileOut;
//		try
//		{
//			fileOut=new FileOutputStream(pr);
//			workbook.write(fileOut);
//			fileOut.close();
//		}
//	catch(Exception e)
//		{
//		e.printStackTrace();
//		}
	}
}

