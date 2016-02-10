/*
 * OSRG Tool - Order Status Report Generator
 * 
 * Written by:  Mike McMahon, A.Sc.T.
 *              mike@3rdgear.ca
 *              647-207-4132
 *
 * Date: September 2, 2015
 * 
 * Version: 1.00.0
 * 
 * Description: This tool is used to convert raw data from a .csv file created by Q2C into a formatted .xls
 *              spreadsheet.
 * 
 * This application consists of four classes:
 * 		- ClientCSV - This is the class that has the main() thread and generates the UI.
 * 		- FormatCSV - Class used to format the raw .csv file.
 * 		- CreateXLS - Class used to create and format the final .xls spreadsheet.
 * 		- Clean - Class used for file clean up.
 *  
 *
 * This class CreateXLS consists of the following components:
 *      - class data fields 
 *      - method: public String convertCSVToXLSX(String workFile, boolean withOptions)
 *      - method: private Map<String, CellStyle> createStyles(Workbook wb)
 *      
 */

import java.util.*;
import java.io.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;


public class CreateXLS 
{
	//class data fields
	private Workbook workBook;
	private Sheet workSheet;
	private String currentLine;
	private int rowNum;
	private int colNum;
	private File xlsFile;
	private BufferedReader reader;
	private FileOutputStream out;
	 
     
    //method converts the formated .csv file referenced by the string workFile to the final .xls 
    //spreadsheet.
	public String convertCSVToXLSX(String workFile, boolean withOptions) throws IOException
	{	
		try
		{
			//temporary .xls file.
			xlsFile = new File("..\\tempXLS.xls");
			
			//create new .xls workbook
			workBook = new HSSFWorkbook();
        
			//map styles to this workbook
			Map<String, CellStyle> styles = createStyles(workBook);
        
			//create a work sheet in the workbook
			workSheet = workBook.createSheet("Order Status Report");
		
			//freeze the top row
			workSheet.createFreezePane(0, 1);
        
			//read in the contents of the .csv file and input them into 
			//the .xls work sheet
			rowNum = 0;
			colNum = 0;
			reader = new BufferedReader(new FileReader(workFile));
        
			while ((currentLine = reader.readLine()) != null) 
			{
				String str[] = currentLine.split(";");
            
				Row currentRow = workSheet.createRow(rowNum);
            
				if (rowNum == 0)
				{
					currentRow.setHeightInPoints(16f);
            	
					for(int i=0; i < str.length; i++)
					{
						Cell cell = currentRow.createCell(i);
						cell.setCellValue(str[i]);
						cell.setCellStyle(styles.get("header")); 
						colNum = i;   
					}
				}
           
				if (rowNum > 0)
				{
					for(int i = 0; i < str.length; i++)
					{
						Cell cell = currentRow.createCell(i);
						cell.setCellValue(str[i]); 
						cell.setCellStyle(styles.get("body"));
					}
				}
				//increment to next record in .csv file
				rowNum++;
			}
        
			//Autosize the columns of the spreadsheet
			for (int i = 0; i < colNum; i++)
			{
				workSheet.autoSizeColumn(i);  	
			}
        
			//Set the Autofilter
			if (withOptions)
			{
				workSheet.setAutoFilter(CellRangeAddress.valueOf("A1:O1"));
			}
			else
			{
				workSheet.setAutoFilter(CellRangeAddress.valueOf("A1:L1"));
			}
        
			//write out the new .xls workbook to a temporary file
			out =  new FileOutputStream(xlsFile);
			workBook.write(out);
			}
			finally
			{
				//close the reader, writer, and .xls workbook
				reader.close();
				out.close();
				workBook.close();
			}
		
        //return the file path to the temporary workbook as a string
        return xlsFile.toString();
	}
	
	
	//method for defining .xls spreadsheet formats and styles
	private Map<String, CellStyle> createStyles(Workbook wb)
	{
		Map<String, CellStyle> styles = new HashMap<String, CellStyle>();
        
		//workbook style 0
        CellStyle style0 = wb.createCellStyle();
        
        //workbook style 1
        CellStyle style1 = wb.createCellStyle();
        
        //style 0 definition
        Font headerFont = wb.createFont();
        headerFont.setFontHeightInPoints((short) 12);
        headerFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
        style0.setAlignment(CellStyle.ALIGN_CENTER);
        style0.setWrapText(true);
        style0.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style0.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style0.setFont(headerFont);
        styles.put("header", style0);
        
        //style 1 definition
        Font font = wb.createFont();
        font.setBoldweight(Font.BOLDWEIGHT_NORMAL);
        style1.setAlignment(CellStyle.ALIGN_LEFT);
        style1.setWrapText(true);
        style1.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        style1.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style1.setFont(font);
        styles.put("body", style1);
        
        return styles;
	}	
}