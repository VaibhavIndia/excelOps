package com.fusion.base;


/* main code 
 * to extract the  value from execl on basis of employee id 
 * and to extract the value from exel on basis of rownumber and column number
 */


import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ReadExcel {

	public static void main(String[] args) throws Exception {

		//String a = getdata("Workplace Culture", "B6");
		String a = getDatafromParticularCell("Workplace Culture", "B6");
		System.out.println(a);
		
	}
	
	
	// Return Array list from particular test case
	public static ArrayList<String> getdata(String sheetName, String testCaseName, String userID) throws Exception
	{
		// TODO Auto-generated method stub
		DataFormatter formatter = new DataFormatter();
		FileInputStream fls = new FileInputStream("C:\\Selenium\\Workspace\\FusionSurvey\\src\\main\\java\\com\\fusion\\testdata\\Fusion Report Builder.xlsx");
		// get the control of excel file 
		XSSFWorkbook wb = new XSSFWorkbook(fls);
		ArrayList<String> a = new ArrayList<String>();
		
		int totalSheet = wb.getNumberOfSheets();
		int rowCount = 0 , columnCount =0, column =0;
		for (int i =0; i<totalSheet;i++)
		{
			if(wb.getSheetName(i).equalsIgnoreCase(sheetName))
				{
					XSSFSheet sheet = wb.getSheetAt(i);
					Iterator<Row> allrows = sheet.iterator();
					while(allrows.hasNext())
					{
						Row row = allrows.next();
						Iterator<Cell> cell = row.cellIterator();
						while(cell.hasNext())
						{
							Cell cellvalue = cell.next();
							if(formatter.formatCellValue(cellvalue).equalsIgnoreCase(testCaseName))
							//if(cellvalue.getStringCellValue().equalsIgnoreCase("Created"))
							{
								column = columnCount;
								System.out.println("yes");
								System.out.println(rowCount);
								System.out.println(columnCount);
								System.exit(0);
							}
							columnCount++;
						}
						rowCount++;
						columnCount=0;
						while(allrows.hasNext())
						{
							Row r = allrows.next();
							if(formatter.formatCellValue(r.getCell(column)).equalsIgnoreCase(userID))
							//if(r.getCell(column).getStringCellValue().equalsIgnoreCase("8"))
							{
								Iterator<Cell> values = r.cellIterator();
								while(values.hasNext())
								{
									//System.out.println("ok");
									Cell c = values.next();
									a.add(formatter.formatCellValue(c));
									//System.out.print(formatter.formatCellValue(c));
									//System.out.println(value);
								}
							}
						}
					}
					
					
				}
				
		}

		//System.out.println(rowCount);
		//System.out.println(columnCount);
		return a;
	}
	

	public static String getdata(String sheetName, String cellID) throws Exception
	{
		// TODO Auto-generated method stub
		DataFormatter formatter = new DataFormatter();
		FileInputStream fls = new FileInputStream("C:\\Selenium\\Workspace\\FusionSurvey\\src\\main\\java\\com\\fusion\\testdata\\Fusion Report Builder.xlsx");
		// get the control of excel file 
		XSSFWorkbook wb = new XSSFWorkbook(fls);
		ArrayList<String> a = new ArrayList<String>();
		String temps="", fstring="";
		int totalSheet = wb.getNumberOfSheets();
		int rowCount = 0 , columnCount =0, column =0;
		for (int i =0; i<totalSheet;i++)
		{
			if(wb.getSheetName(i).equalsIgnoreCase(sheetName))
				{
					XSSFSheet sheet = wb.getSheetAt(i);
		
					FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
					// suppose your formula is in B3
					CellReference cellReference = new CellReference(cellID); 
					Row row = sheet.getRow(cellReference.getRow());
					Cell cell = row.getCell(cellReference.getCol()); 

					if(cell!=null)
					{
						switch (evaluator.evaluateInCell(cell).getCellType()) 
						{
						    case Cell.CELL_TYPE_BOOLEAN:
						    	boolean tempb=true;
						    	temps = Boolean.toString(tempb);
						    	fstring = temps;
						    	//System.out.println(temps);
						    	
						        //System.out.println(cell.getBooleanCellValue());
						        break;
						    case Cell.CELL_TYPE_NUMERIC:
						    	if(cell.getNumericCellValue()<1) 
						    	{
						    		int tempi = 0;
						    		tempi = (int) Math.round(cell.getNumericCellValue()*100);
						    		//System.out.println(tempi);
						    		temps = Integer.toString(tempi);
						    		//System.out.println(temps);
						    		fstring =temps;
						    		//System.out.println((cell.getNumericCellValue())*100);
						    	}
						    	//System.out.println(temps +"sad");
						    	break;
						        
						    case Cell.CELL_TYPE_STRING:
						        temps= cell.getStringCellValue();
						        fstring = temps ;
						    	
						        //System.out.println(temps + "111111");
						        //System.out.println(cell.getStringCellValue());
						        break;
						    case Cell.CELL_TYPE_BLANK:
						        break;
						    case Cell.CELL_TYPE_ERROR:
						        break;
						
						    // CELL_TYPE_FORMULA will never happen
						    case Cell.CELL_TYPE_FORMULA: 
						        break;
						}
				}
						
			}
		}
		return fstring;
	}
	
	
	// get data from excel sheet by providing cell number (like F12) and sheetname
		public static String getDatafromParticularCell(String sheetName,String cellName) throws IOException{
			FileInputStream fls = new FileInputStream("D:\\Fusion culture survey\\Docs\\Questions.xlsx");
			
			// get the control of excel file 
			XSSFWorkbook wb = new XSSFWorkbook(fls);
			int totalSheets = wb.getNumberOfSheets();
			for(int i =0; i<totalSheets;i++)
			{
				if(wb.getSheetName(i).equalsIgnoreCase(sheetName))
				{
					XSSFSheet sheet = wb.getSheetAt(i);
					Pattern r = Pattern.compile("^([A-Z]+)([0-9]+)$");
				    Matcher m = r.matcher(cellName);
				    if(m.matches())
				    {
				    	 String columnName = m.group(1);
				    	 int rowNumber = Integer.parseInt(m.group(2));
				    	 if(rowNumber > 0) {
				    		 
				    		 return sheet.getRow(rowNumber-1).getCell(CellReference.convertColStringToIndex(columnName)).getStringCellValue();
				    		 //xyz= sheet.getRow(rowNumber-1).getCell(CellReference.convertColStringToIndex(columnName));
				    		 
				             //return xyz;
				         }
				    }
				}
				
			}
			
		    
		    return null;
		}
	
}
	

