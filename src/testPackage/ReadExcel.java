package testPackage;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;


import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/*
 * Test Class to Read Excel Files
 */

public class ReadExcel {

	/**
	 * @param args
	 * @throws IOException 
	 * @throws OpenXML4JException 
	 */
	public static void main(String[] args){
		// TODO Auto-generated method stub
		
		try {
			//InputStream inp = new FileInputStream("/Users/americosavinon/Documents/workspace/ApachePOI/TestExcel.xlsx");
			InputStream inp = new FileInputStream("/Users/americosavinon/Documents/workspace/ApachePOI/Hair Care IBBG - rev 11-3-11.xlsx");
			Workbook wb1 = new XSSFWorkbook(inp);
			
			//test_SimpleRead( wb1 );
			
			test_ReadNamedRange( wb1);
			
			  
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}	
		catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	public static void test_ReadNamedRange( Workbook wb ){
		
		try{
			String cname = "PricingClub";

		    // retrieve the named range
		    int namedCellIdx = wb.getNameIndex(cname);
		    Name aNamedCell = wb.getNameAt(namedCellIdx);

		    // retrieve the cell at the named range and test its contents
		    AreaReference aref = new AreaReference(aNamedCell.getRefersToFormula());
		    CellReference[] crefs = aref.getAllReferencedCells();
		    
		    //Row number
		    int current_row = 0;
		    StringBuffer sb = new StringBuffer();
		    sb.append("<table><tr>");
		    //Cell contents
		    String current_content = "";
		    
		    for (int i=0; i<crefs.length; i++) {
		        Sheet s = wb.getSheet(crefs[i].getSheetName());
		        Row r = s.getRow(crefs[i].getRow());
		        Cell c = r.getCell(crefs[i].getCol());
		        
		        //Cell Contents
		        current_content = ""+ c;
		        
		        //Check the current row number: if this value is different from the last one, it means we are on a new row.
		        if( current_row != crefs[i].getRow()){
		        	current_row = crefs[i].getRow();
		        	
		        	sb.append("</tr>");
		        	sb.append("<tr>");
		        }
		        sb.append("<td>"+ current_content +"</td>");
		        
		        //System.out.println("Row number:" + crefs[i].getRow() + " :: Cell content:" + c);
		        
		    }
		    sb.append("</tr></table>");
		    
		    System.out.println( sb.toString());
		}
		catch(IllegalArgumentException e){
			System.out.println("ERROR IN NAME RANGE, maybe the name range does not exist, check the name");
			e.printStackTrace();
		}
		
	}
	
	
	public static void test_SimpleRead(Workbook wb ){
		Sheet sheet = wb.getSheetAt(0);  
		  
		int rows; // No of rows  
		rows = sheet.getPhysicalNumberOfRows();  
		      
		System.out.println("the total number of rows are:"+rows);
	}
}
