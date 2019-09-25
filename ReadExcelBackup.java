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


public class ReadExcelBackup {

	public static void main(String[] args) throws Exception {
		

		
//		for(int n=0;n<35;n++)
//		{
//			String number = Integer.toString(n);
//		ArrayList<String> a =  getdata("Raw Data", "ID", number);
//		for(String a1 : a )
//		{
//			System.out.println(a1);
//		}
//		System.out.println("above data is for employee "+ number);
//		//System.out.println(a.get(46));
//		}
//		
//		String data = getDataFromRequiredRowAndCell("Raw Data", 5, 3);
//		System.out.println(data);
//		
////		
		XSSFCell c = getCell("Page 3", "B32");
		System.out.println(c);
//		
//		String na = apachemethod("Page 3", "F10");
//		System.out.println(na);
//		
//		String a = getCellValueAsString("Page 3", "F10");
//		System.out.println(a);
		
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
	
	// Get data from required cell by using row number and column number
	public static String getDataFromRequiredRowAndCell(String sheetName, int rowNumber, int columnNumber ) throws IOException
	{
		DataFormatter formatter = new DataFormatter();
		FileInputStream fls = new FileInputStream("C:\\Selenium\\Workspace\\FusionSurvey\\src\\main\\java\\com\\fusion\\testdata\\Fusion Report Builder.xlsx");
		// get the control of excel file 
		XSSFWorkbook wb = new XSSFWorkbook(fls);
		String data="";
		int totalCountOfSheets =wb.getNumberOfSheets();
		for(int i=0;i<totalCountOfSheets;i++)
		{
			if(wb.getSheetName(i).equalsIgnoreCase(sheetName))
			{
				XSSFSheet mySheet = wb.getSheetAt(i);
				Row r = mySheet.getRow(rowNumber);
				Cell cell = r.getCell(columnNumber);
				data = formatter.formatCellValue(cell);
				data = cell.getStringCellValue();
				System.out.println(cell);
			}
		}
		
		
		return data;
	}
	
	// get data from excel sheet by providing cell number (like F12) and sheetname
	public static XSSFCell getCell(String sheetName,String cellName) throws IOException{
		FileInputStream fls = new FileInputStream("C:\\Selenium\\Workspace\\FusionSurvey\\src\\main\\java\\com\\fusion\\testdata\\Fusion Report Builder.xlsx");
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
			             return sheet.getRow(rowNumber-1).getCell(CellReference.convertColStringToIndex(columnName));
			         }
			    }
			}
			
		}
		
	    
	    return null;
	}
	
	public static XSSFCell getCellforFormula(String sheetName,String cellName) throws IOException{
		FileInputStream fls = new FileInputStream("C:\\Selenium\\Workspace\\FusionSurvey\\src\\main\\java\\com\\fusion\\testdata\\Fusion Report Builder.xlsx");
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
			             if(sheet.getRow(rowNumber-1).getCell(CellReference.convertColStringToIndex(columnName)).getCellType()==Cell.CELL_TYPE_FORMULA)
			             {
			            	 //print the formula
			            	 //System.out.println("Formula is " + cell.getCellFormula());
			            	 
			             }
			         }
			    }
			}
			
		}
		
	    
	    return null;
	}

	public static String apachemethod(String sheetName, String columnName) throws IOException
	{
		FileInputStream fls = new FileInputStream("C:\\Selenium\\Workspace\\FusionSurvey\\src\\main\\java\\com\\fusion\\testdata\\Fusion Report Builder.xlsx");
		// get the control of excel file 
		XSSFWorkbook wb = new XSSFWorkbook(fls);
		String value ="";
		int totalSheets = wb.getNumberOfSheets();
		for(int i =0; i<totalSheets;i++)
		{
			if(wb.getSheetName(i).equalsIgnoreCase(sheetName))
			{
				XSSFSheet sheet = wb.getSheetAt(i);
				FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
				
				// suppose your formula is in B3
				CellReference cellReference = new CellReference(columnName); 
				Row row = sheet.getRow(cellReference.getRow());
				Cell cell = row.getCell(cellReference.getCol()); 
				
				CellValue cellValue = evaluator.evaluate(cell);
				
				switch (cellValue.getCellType())
					{
				    case Cell.CELL_TYPE_BOOLEAN:
				    	boolean tempb = cellValue.getBooleanValue();
				    	value = Boolean.toString(tempb);
				    	//System.out.println(cellValue.getBooleanValue());
				        break;
				    case Cell.CELL_TYPE_NUMERIC:
				        int tempi = (int) cellValue.getNumberValue();
				        int tempite=0;
				        if(tempi<1)
				        {
				        	String t = Integer.toString(tempi);
				        	Double d = Double.parseDouble(t) * 100;
				        	tempite  = (int) Math.round(d);
				        }
				        value = Integer.toString(tempite);
				    	
				    	//System.out.println(cellValue.getNumberValue());
				        break;
				    case Cell.CELL_TYPE_STRING:
				    	value = cellValue.getStringValue();
				    	
				        //System.out.println(cellValue.getStringValue());
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
		return value;
	}
	
	
	
	
	/* This method for the type of data in the cell, extracts the data and
    * returns it as a string.
    */
   public static String getCellValueAsString(String sheetName, String columnName) throws IOException 
   {
	   
		FileInputStream fls = new FileInputStream("C:\\Selenium\\Workspace\\FusionSurvey\\src\\main\\java\\com\\fusion\\testdata\\Fusion Report Builder.xlsx");
		// get the control of excel file 
		XSSFWorkbook wb = new XSSFWorkbook(fls);
		String strCellValue = null;
		int totalSheets = wb.getNumberOfSheets();
		for(int i =0; i<totalSheets;i++)
		{
			if(wb.getSheetName(i).equalsIgnoreCase(sheetName))
			{
				XSSFSheet sheet = wb.getSheetAt(i);
				FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
				
				// suppose your formula is in B3
				CellReference cellReference = new CellReference(columnName); 
				Row row = sheet.getRow(cellReference.getRow());
				Cell cell = row.getCell(cellReference.getCol()); 
				
				
				if (cell != null) 
				{
					switch (cell.getCellType()) 
					{
						case Cell.CELL_TYPE_STRING:
						strCellValue = cell.toString();
						break;
						
						case Cell.CELL_TYPE_NUMERIC:
							if (DateUtil.isCellDateFormatted(cell)) 
							{
								SimpleDateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
								strCellValue = dateFormat.format(cell.getDateCellValue());
							}
							else 
							{
								Double value = cell.getNumericCellValue();
								Long longValue = value.longValue();
								strCellValue = new String(longValue.toString());
							}
							break;
           
						case Cell.CELL_TYPE_BOOLEAN:
							strCellValue = new String(new Boolean(
							cell.getBooleanCellValue()).toString());
							break;
           
						case Cell.CELL_TYPE_BLANK:
							strCellValue = "";
							break;
					}
				}
				
			}
			
       }
		return strCellValue;
   }
   


}
	

