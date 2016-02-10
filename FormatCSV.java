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
 * This class FormatCSV consists of the following components:
 *      - class data fields   
 *      - method: public String formatedCSV(File file, boolean withOptions)
 *      
 */

import java.util.*;
import java.io.*;
import org.apache.commons.csv.*;


public class FormatCSV 
{
	//class data fields 
	private File tempFile;
	private CSVParser parser;
	private CSVPrinter printer;
	private List<CSVRecord> recordList;
	private CSVRecord data;
	private List<String[]> headerList;
	
    
    //method converts raw .csv file produced by Q2C to a formated .csv file
    //that can then be converted to the final .xls file.
	public String formatedCSV(File file, boolean withOptions) throws IOException
	{
		try
		{
			//new temporary file
			tempFile = new File("..\\temp.csv");
		
			//create the parser for the .csv file based on the file selected by the user.
			parser = new CSVParser(new FileReader(file), CSVFormat.DEFAULT.withHeader());
		
			//create a printer for generating a temporary/intermediate .csv file.
			//use ';' as the delimiter in the the temp file and not a ',' to avoid
			//issues when aligning columns with ',' in the data.
			printer = new CSVPrinter(new FileWriter(tempFile), CSVFormat.DEFAULT.withDelimiter(';'));
		
			//read all the .csv file records into memory.
			recordList = parser.getRecords();
		
			//Iterate over the collection recordList and remove all "Cancelled" line items.
			for (Iterator<CSVRecord> iterator = recordList.iterator(); iterator.hasNext();)
			{
				data = iterator.next();
				String dataString = data.get("prgrs_pnt");
				if (dataString.contains("Cancelled"))
				{
					iterator.remove();
				}
			}
		
			//Iterator over the collection recordList and remove all "PERU" line items that are 
			//not sub line .
			for (Iterator<CSVRecord> iterator = recordList.iterator(); iterator.hasNext();)
			{
				CSVRecord data = iterator.next();
				String subLine = data.get("sub_ln_itm_id");
				String plant = data.get("shiploc");
				String schedID = data.get("shpschd_id");
				if (plant.contains("CANADA RESALE") && subLine.contains("0") && schedID.contains("2"))
				{
					iterator.remove();
				}
			}
		
			//if 'withOptions' is true, create a formatted .csv file with the optional columns.
			if (withOptions)
			{
				//create temporary .csv file header
				headerList = new ArrayList<String[]>();
				headerList.add(new String[]{"Line Number", "Sub-Line Number", "Unfilled Qty", "Shipped Qty", "Catalog Number", "Catalog Description", "Designation",
						"Shipping Location", "Commit Ship Date", "Not Before Date", "Current Ship Date", "Customer Requested Delivery Date", "Progress Point", 
						"Date Shipped", "Hold Y/N"});
			
				//print the temporary header
				printer.printRecords(headerList);
		
				//get each .csv file record from memory and append to temporary .csv file
				for (CSVRecord record : recordList)
				{
					printer.printRecord(record.get("ln_dsply_seq_nbr"), record.get("sub_ln_itm_id"), record.get("unfld_ord_qty"), record.get("shpd_qty"), record.get("catlg_nbr"), 
							record.get("catlg_desc"), record.get("desnat_desc"), record.get("shiploc"), record.get("commit_shpschd_dt"), record.get("not_b4_dt"), record.get("prom_dt"), 
							record.get("cust_rqst_dlvry_dt"), record.get("prgrs_pnt"), record.get("ship_dt"), record.get("actn_status"));
				}
			}
			//else create a standard report without the optional columns.
			else
			{
				//create temporary .csv file header
				headerList = new ArrayList<String[]>();
				headerList.add(new String[]{"Line Number", "Sub-Line Number", "Unfilled Qty", "Shipped Qty", "Catalog Number", "Catalog Description", "Designation",
						"Shipping Location", "Current Ship Date", "Progress Point", "Date Shipped", "Hold Y/N"});
				
				//print the temporary header
				printer.printRecords(headerList);
		
				//get each .csv file record from memory and append to temporary .csv file
				for (CSVRecord record : recordList)
				{
					printer.printRecord(record.get("ln_dsply_seq_nbr"), record.get("sub_ln_itm_id"), record.get("unfld_ord_qty"), record.get("shpd_qty"), record.get("catlg_nbr"), 
							record.get("catlg_desc"), record.get("desnat_desc"), record.get("shiploc"), record.get("prom_dt"), record.get("prgrs_pnt"),
							record.get("ship_dt"), record.get("actn_status"));
				}
			}
		}
		finally
		{
			//close both the parser and the printer.
			parser.close();
			printer.close();
		}
		
		//return temorary .csv file path as string
		return tempFile.toString();
	}
}
