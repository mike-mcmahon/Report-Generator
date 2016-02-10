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
 * This class Clean consists of the following components: 
 *      - method: public void cleanUp(String filePath)
 *      
 */

import java.io.*;


public class Clean 
{
	public void cleanUp(String filePath) throws FileNotFoundException, IOException
	{
		//delete temporary files used in this application
        //deleted files are typically temporary/intermediate .csv and .xls files
		File file = new File(filePath);
		file.delete();
	}
}
