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
 * This class ClientCSV consists of the following components:
 *      - class data fields
 *      - constructor: public ClientCSV() 
 *          with inner classes: class MyItemListener implements ItemListener,
 *              class MyActionListenerCreate implements ActionListener, and 
 *              class MyActionListenerPick implements ActionListener   
 *      - method: public static void main(String[] args)
 *      
 */

import java.awt.*;
import java.awt.event.*;
import javax.swing.*;
import java.io.*;


public class ClientCSV extends JPanel
{
	//class data fields
    private static final long serialVersionUID = 1L;
	
	//UI components
	private JCheckBox optionalReport;
    private JButton create;
    private JLabel appName;
    private JButton pickFileButton;
    private JTextArea log;
    private JFileChooser fileChooser;
    
    //FormatCSV and CreateXLS objects
    private FormatCSV csvFormater;
    private CreateXLS xlsGenerator;
    private Clean cleaner;
    
    //general data fields
    private int returnVal;
    private File csvFile;
    private String temporaryCSV;
    private String temporaryXLS;
    private boolean fileSelected;
    private boolean withOptions;
    
    
    //UI Constructor
    public ClientCSV()
    {
    	//set container-panel size
        this.setPreferredSize(new Dimension(350, 200));
        
        //create layout manager object and set automatic gap insertion
        GroupLayout layout = new GroupLayout(this);
        this.setLayout(layout);
        layout.setAutoCreateGaps(true);
        layout.setAutoCreateContainerGaps(true);
        
        //create UI components
        optionalReport = new JCheckBox("with optional columns");
        create = new JButton("Create Report");
        appName = new JLabel("Order Status Report");
        pickFileButton = new JButton("Select CSV File");
        log = new JTextArea(5, 20);
        
        //create the component groups for the layout
        layout.setHorizontalGroup(
            layout.createParallelGroup(GroupLayout.Alignment.CENTER)
                .addComponent(pickFileButton)
                .addComponent(log)
                .addComponent(appName)
            .addGroup(layout.createSequentialGroup()
                .addComponent(create)
                .addComponent(optionalReport))                
        );
        
        layout.setVerticalGroup(
            layout.createSequentialGroup()
                .addComponent(appName)
                .addComponent(pickFileButton)
                .addComponent(log)
                .addGroup(layout.createParallelGroup()
                    .addComponent(create)
                    .addComponent(optionalReport))
        );
        
        //create objects from classes Clean, CreateXLS, and FormatCSV
        csvFormater = new FormatCSV();
        xlsGenerator = new CreateXLS();
        cleaner = new Clean();
        
        //action listener for 'select file' JButton
        class MyActionListenerPick implements ActionListener
        {
            public void actionPerformed(ActionEvent ae)
            {
                fileChooser = new JFileChooser();
                fileSelected = false;
                
                if (ae.getSource() == pickFileButton)
                {
                    returnVal = fileChooser.showOpenDialog(ClientCSV.this);
                        
                    if (returnVal == JFileChooser.APPROVE_OPTION)
                        {
                            csvFile = fileChooser.getSelectedFile();

                            log.append("Opening: " + csvFile.getName() + "." + "\n");
                            
                            fileSelected = true;
                        }
                        else 
                        {
                            log.append("Opening command cancelled by user." + "\n");
                            log.append("No file selected." + "\n");
                        }
                }
                
            }  
        }
        
        //action listener for JCheckbox
        class MyItemListener implements ItemListener
        {
            public void itemStateChanged(ItemEvent ie)
            {
                //boolean value must toggle between true and false
                int state = ie.getStateChange();
                if (state == ItemEvent.SELECTED)
                {
                    withOptions = true;
                }
                else
                {
                    withOptions = false;
                }
            }
        }
        
        //action listener for 'create' JButton
        class MyActionListenerCreate implements ActionListener
        {
            public void actionPerformed(ActionEvent ae)
            {
                //ensure a file is selected prior to executing report generation
                if (ae.getSource() == create && fileSelected)
                {
                	try
                	{
                	//create temporary formated .csv file
                    temporaryCSV = csvFormater.formatedCSV(csvFile, withOptions);
                    
                    //create .xls file
                    temporaryXLS = xlsGenerator.convertCSVToXLSX(temporaryCSV, withOptions);
                    log.append(".xls file is now created" + "\n");
                    
                    
                    //from the 'temporaryXLS string, create a new xls file.
                    File xlsFile = new File (temporaryXLS);
             
                    //blank out file to save dialog before opening file save dialog box
                    
                    //open a save file dialog box to save .xls file to user defined location
                    returnVal = fileChooser.showSaveDialog(ClientCSV.this);
                    if (returnVal == JFileChooser.APPROVE_OPTION)
                    {
                    	//create a file using the save dialog filechooser
                    	File saveFile = fileChooser.getSelectedFile();
                    	xlsFile.renameTo(saveFile);
                    	
                    	log.append("Saving: " + saveFile.getPath() + "\n");
                    	
                    	//allow for multiple reports to be generated with one instance of this application
                        fileSelected = false;
                    }
                    else
                    {
                    	log.append("Save command cancelled by user." + "\n");
                    }
                     
                    //clean up temporary files
                    cleaner.cleanUp(temporaryCSV);
                    cleaner.cleanUp(temporaryXLS);
                	}
                	catch (IOException e)
                	{
                		e.printStackTrace();
                		System.out.println(e.getMessage());
                	}
                }
            }
        }
        
        //instantiate the action listeners
        pickFileButton.addActionListener(new MyActionListenerPick());
        optionalReport.addItemListener(new MyItemListener());
        create.addActionListener(new MyActionListenerCreate());
    }
    
    
    //main thread
    public static void main(String[] args)
    {
    	JFrame frame = new JFrame("OSRG Tool");
    	frame.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
    	frame.add(new ClientCSV());
    	frame.pack();
    	frame.setVisible(true);
    }
}
