﻿using Microsoft.Win32;
using Ookii.Dialogs.WinForms;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Vbe.Interop;
using Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Data;
using ExcelLibrary;
using Siemens.Engineering;
using Siemens.Engineering.Compiler;
using Siemens.Engineering.Hmi;
using Siemens.Engineering.HW;
using Siemens.Engineering.HW.Features;
using Siemens.Engineering.HW.Utilities;
using Siemens.Engineering.Library;
using Siemens.Engineering.Settings;
using Siemens.Engineering.SW;
using Siemens.Engineering.SW.Blocks;
using Siemens.Engineering.SW.Types;


namespace EPLAN_TIA
{
    
    public partial class Form1: Form
    { 
        public Form1()
        {
            InitializeComponent();
            AppDomain CurrentDomain = AppDomain.CurrentDomain;
            CurrentDomain.AssemblyResolve += new ResolveEventHandler(MyResolver);
        }

        //Security variables 

        int copy = 0;

        //Excel variables

        //Excel excel = new Excel()

        Excel.Application xlApp/*= new Excel.Application()*/; 
        Excel.Application xlApp_2/*= new Excel.Application()*/; 
        Excel.Application xlApp_3/*= new Excel.Application()*/; 
        Excel.Application xlApp_4 /*= new Excel.Application()*/; 

        object misValue = System.Reflection.Missing.Value;

        //SPSData
        
        Excel.Workbook xlWorkBook ;
        Excel.Worksheet xlWorkSheet ;

        //SPS Verbindungen

        Excel.Workbook xlWorkBook_2;
        Excel.Worksheet xlWorkSheet_2;

        //SPS neue Verbindungen

        Excel.Workbook xlWorkBook_3;
        Excel.Worksheet xlWorkSheet_3;

        //SPS TIA

        Excel.Workbook xlWorkBook_4;
        Excel.Worksheet xlWorkSheet_4;
        
        //TIA Variables 

        public static Project project { get; set; }
        public TiaPortal MyTiaPortal { get; set; }
        public Project MyProject { get; set; }

        //Language selected (English by default)
        string languageSelected = "EN";

        //Status variables
        string errorMessage;

        //Strings for messages
        string executing;
        string invalidPath;
        string programmFinished;
        string programmStopped;
        string nothingFound;
        string longNamesTextBeggin;
        string longNamesTextEnd;

        //English variables
        string errorMessageEN = "Error: ";

        string executingEN= "Programm is in execution...";

        string programmFinishedEN= "The programm has finished";
        string programmStoppedEN="The programm has been stopped due to an error: there are variables with names longer than 15" +
            " characters. Please fix it and try again.";

        string nothingFoundEN = "No information was found in the Excel. Please check it and try again.";


        string txtPathField1EN = "Select path where the Excel file PLC DATA (EPLAN) is located";
        string txtPathField2EN = "Select path where the Excel file Data Conexions (EPLAN) is located";
        string txtPathField3EN = "Select path where the Excel file export (TIA) is located";
        string txtPathField4EN = "Select older to save the data backups";

        string browseEN = "Browse";

        string statusEN="Status";
        string invalidPathEN = "Invalid path";

        //new English variables 

        string closeAllInstances = "All TIA Portal instances will be closed. Do you want to continue?";
        string closeAllInstancesWarning = "Close all TIA Portal instances";

        //German variables
        string errorMessageDE = "Fehler: ";

        string executingDE= "Programm wird ausgeführt...";

        string programmFinishedDE= "Das Programm ist beendet";

        string programmStoppedDE = "Das Programm wurde aufgrund eines Fehlers gestoppt: es gibt Variablen mit " +
            "Namen, die länger als 15 Zeichen sind. Bitte beheben Sie den Fehler und versuchen Sie es erneut.";

        string nothingFoundDE = "Keine Information wurde in der Excel-Datei gefunden. Bitte überprüfen Sie es und versuchen Sie es erneut.";

        string longNamesTextBegginDE = "Die folgenden Variablen haben einen Namen länger als 15 Zeichen: \n \n";
        string longNamesTextEndDE = "\nBitte beheben Sie den Fehler und versuchen Sie es erneut.";

        string txtPathField1DE = "Wähle den Pfad, in dem sich die Excel SPS aus EPLAN befindet";
        string txtPathField2DE = "Wähle den Pfad, in dem sich die Excel Verbindungen aus EPLAN befindet";
        string txtPathField3DE = "Wähle den Pfad, in dem sich die Excel Datei Export aus TIA befindet";
        string txtPathField4DE = "Wählen Sie den Ordner zum Speichern der Sicherungen Datein";

        string browseDE = "Durchsuchen";

        string statusDE = "Status";
        string invalidPathDE = "Ungültiger Pfad";

        //Lists

        List<InformationPLC> informationSPs = new List<InformationPLC>(); //keep all the information of the differents SPS avaible
        List<conexionEPLAN> dataNotSimilar = new List<conexionEPLAN>();

        //Class for the information of a PLC
        public class InformationPLC 
        {
            public string serialNum { get; set; }
            public string eplanName { get; set; }
            public string IP { get; set; }                                
            public string cpuPPAL { get; set; }                          
            public string startAdresse { get; set; }                    
            public List<SW_HW> adresseSW_HW = new List<SW_HW>(); 
            
            public InformationPLC (string serianum, string eplanname, string ip, string cpuppal, string startadress, List <SW_HW> adresse_sw_hw)
            {
                serialNum = serianum;
                eplanName = eplanname;
                IP = ip;
                cpuPPAL = cpuppal;
                startAdresse = startadress;
                adresseSW_HW = adresse_sw_hw; 
            }
        }


        public class SW_HW
        {
            public string swAd { get; set; }
            public string hwAd { get; set; }

            public SW_HW (string swad, string hwad)
            {
                swAd = swad;  //keep the values of the software (adress) 
                hwAd = hwad;    //keep the values of the hardware (adress)
            }

        }


        //Data to import in TIA 

        string projectName; 
       

        //Load form
        private void Form1_Load(object sender, EventArgs e)
        {
            
            //Set visualisation to normal size
            WindowState = FormWindowState.Normal;
            //Load texts in english by default

            if (languageSelected == "DE")
            {
                languageDE();
            }
            else
            {
                languageEN();
            }


        }
       
        //Start the process
        public void btn_Start_Click(object sender, EventArgs e)
        {
            xlApp = new Excel.Application(); 
            xlWorkBook = xlApp.Workbooks.Open(txt_Path1.Text); 
            xlWorkSheet = xlWorkBook.Worksheets.get_Item(1);

            xlApp_2 = new Excel.Application();
            xlWorkBook_2 = xlApp_2.Workbooks.Open(txt_Path3.Text); 
            xlWorkSheet_2 = xlWorkBook_2.Worksheets.get_Item(1);

            xlApp_3 = new Excel.Application();
            xlWorkBook_3 = xlApp_3.Workbooks.Add();
            xlWorkSheet_3 = xlApp_3.Worksheets.Add();


            xlApp_4=new Excel.Application();
            xlWorkBook_4= xlApp_4.Workbooks.Open(txt_Path4.Text); 
            xlWorkSheet_4 = xlWorkBook_4.Worksheets.get_Item(1); 



            int startColumn =1; //in Excel DateiSPS must be the first column. In opposite in the excel Verbinden must be the second column. 
            int startColumn_2 = 9; //in Excel Verbindungen the first value ist for the source and the next for the destination
            
            //Show message

            txt_Status.Text = executing;

            //Excel variable for the new file (Excel Verbindungen 2) 

            
            //read the data from "SPSDatei". 

            getinfoPLC(startColumn);

            txt_Status.Text = "The information of the Excel SPS Datei was correct readed";

            //read the second excel but only with the data from the SPS of the previous data. 

            getinfConexion(startColumn_2,informationSPs,xlWorkSheet_3);

            txt_Status.Text = "The information of the Excel SPS Verbindungen was correct readed";

            //Write the relevant information in other Excel. 

            txt_Status.Text = "Exporting information to the new Excel";


            SaveCloseExcel(txt_Path1.Text, xlApp, xlWorkBook,xlWorkSheet, misValue); //save the SPSDATEI
            SaveCloseExcel(txt_Path3.Text, xlApp_2, xlWorkBook_2, xlWorkSheet_2, misValue); //save the SPS conexions
            

            //the third Excel keep opened and the excel from TIA will be opened, and compared with the other. 

            txt_Status.Text = "Comparing data between both programs";

            //Compare data.

            CompareEPLANTIA(xlWorkSheet_3);


            //Close the excel 3 and excel 4. 

            string path = txt_Path2.Text;
            string nameWork = "EPLAN_Datei.xlsx";
            string exportFilename = @"" + path + @"\" + nameWork;
             

            if(File.Exists(exportFilename))
            {
                do
                {
                    copy++; //Increment the copy               
                    nameWork = "EPLAN_Datei" + copy + ".xlsx";
                    exportFilename = @"" + path + @"\" + nameWork;
                } while (File.Exists(exportFilename)); 
                
            }

            xlWorkBook_3.SaveAs(exportFilename);
            Marshal.ReleaseComObject(xlWorkBook_3);
            Marshal.ReleaseComObject(xlWorkSheet_3);
            Marshal.ReleaseComObject(xlApp_3);
            KillSpecificExcelFileProcess(nameWork);

            //Save the excel. 

            SaveCloseExcel(txt_Path4.Text, xlApp_4, xlWorkBook_4,xlWorkSheet_4, misValue);

            // Import the data in TIA. 

                //Start Instance in TIA Portal.
                StartTIA();
                
                //Create a new Project. 

                CreateProject(); 

                //Read the structure in informationSPs to import in TIA. 

                
                foreach(InformationPLC device in informationSPs)
                {

                    AddDisp(device.serialNum, device.eplanName, device.cpuPPAL, project); 
                    
                }
                



            //Show message.

            txt_Status.Text = programmFinished;


            //Kill the background opened processes. 

            GC.Collect();
            GC.WaitForPendingFinalizers();



        }

        //--Functions for the first part of the APP. (Obtein the data of the EPLAN) 

        public void getinfoPLC(int startColumn) 
        {
            int cpuPPALES=0; //count the number of CPU with IP. 
            string serialNum;
            string eplanMark;
            string IP;
            string cpuPPAL;
            string startAdresse;
            string swAd;
            string hwAd;
            List<SW_HW> adresseSW_HW1;      
            bool finished = false;
            bool finished1= false;
            int line = 2;       
            string anzahlText;   
            int colRead = 15;   
            int iList = 1;
            int cutString = 0; 
            //Repeat the process until an empty cell oof number of this geaete is found
            do
            {
                //try
                //{
                   adresseSW_HW1 = new List<SW_HW>(); //reload a new direction in memory for the next values. 

                    serialNum = (string)(xlWorkSheet.Cells[line, 1] as Excel.Range).Value;
                    anzahlText = (string)(xlWorkSheet.Cells[line, 2] as Excel.Range).Value;

                    if (serialNum != null) //Si es un geräte util de importar en TIA
                    {
                        //Save the easy data, because its only in each column of the excel

                        eplanMark = (string)(xlWorkSheet.Cells[line, 4] as Excel.Range).Value;
                        IP = (string)(xlWorkSheet.Cells[line, 10] as Excel.Range).Value;

                        if (IP != null)
                        {
                            cpuPPALES++; //we have other CPU that is central CPU.
                        }
                        
                        cpuPPAL = (string)(xlWorkSheet.Cells[line, 11] as Excel.Range).Value;
                        cutString = cpuPPAL.IndexOf("=");
                        cpuPPAL=cpuPPAL.Substring(cutString+1);
                        cutString = cpuPPAL.IndexOf("+"); //reuse the variable 
                        projectName = cpuPPAL.Substring(0, cutString);
                        cutString = eplanMark.IndexOf("+");     
                        eplanMark = eplanMark.Substring(cutString+1); //take only the name of each device. 
                        startAdresse = (string)(xlWorkSheet.Cells[line, 13] as Excel.Range).Value;

                        do
                        {
                            //--Read the SWpin data

                            SW_HW swAdhwAd1 = new SW_HW("null", "null"); //create a new linked list. 

                            swAd = (string)(xlWorkSheet.Cells[line, colRead] as Excel.Range).Value;
                            hwAd = (string)(xlWorkSheet.Cells[line, (colRead + 1)] as Excel.Range).Value;

                            //If it was an empty cell, the check process for the SWAdresses is over

                            if (String.IsNullOrEmpty(swAd) || String.IsNullOrEmpty(hwAd)) //if one of the columns are empty, is not a correct par value
                            {
                                colRead = 15; //reset for the next iteration
                                finished = true;
                            }

                            else

                            {
                                swAdhwAd1.swAd = swAd;
                                swAdhwAd1.hwAd = hwAd;
                                adresseSW_HW1.Add(swAdhwAd1); //at the end, there will be a huge dictionary with the SWAdress and the HWAdress for each serialNummer.                               
                                colRead = colRead + 2; //the next column
                                finished = false;
                            }



                        } while (!finished);

                        InformationPLC informationSPS = new InformationPLC(serialNum, eplanMark, IP, cpuPPAL, startAdresse, adresseSW_HW1);

                        informationSPs.Capacity = iList; //increment the size of the list dinamic.
                        informationSPS.serialNum = serialNum;
                        informationSPS.eplanName = eplanMark.Replace('=', ' ');
                        informationSPS.IP = IP;
                        informationSPS.cpuPPAL = cpuPPAL;
                        informationSPS.startAdresse = startAdresse;
                        informationSPS.adresseSW_HW = adresseSW_HW1; 
                        informationSPs.Add(informationSPS); //Assing the values

                        iList++;
                        line++;
                         
                    }

                    else if (anzahlText!= null && serialNum == null) //we have a geräte that is not possible to import in TIA. 
                    {
                        line++; 
                    }
                    else if (anzahlText ==null) //we have read all the gerätes.
                    {
                        finished1=true;
                        //line = 2; //reset the value
                        //iList = 1; //reset the value
                        txt_Status.Text = " No more data to read in SPS Datei";
                    }
                    

                //}
                //If the column of the SPS-Typ is empty
                //catch
                //{
                //    //Controll if there is more SPS Data

                //    anzahlText = (string)(xlWorkSheet.Cells[line, 2] as Excel.Range).Value;

                //   if (anzahlText ==null)
                //    {
                //        finished1=true; //there are no more SPS data. 
                //    }

                //    else
                //    {
                //        line++; //read the next line
                //    }
                //}

                txt_Status.Text = " Reading the data of the SPS";

            } while (!finished1);

        }

        public void getinfConexion(int startColumn, List<InformationPLC> informationSPs,Excel.Worksheet xlWorksheet )
        {
            int line = 2;  //read from the second line. 
            string textDestination;
            string textSource;
            string name;
            string hdwadd;
            string sfadd;
            string senactadd;
            int positionSen;
            int positionSenFin; 
            int position; 
            int newLine =1; //write from the first line. 
            bool finished=false;
          

            xlWorksheet.Activate();
       
            foreach (InformationPLC item in informationSPs)
            {   
                do
                {
                    name = item.eplanName;
                    textSource = (string)(xlWorkSheet_2.Cells[line, startColumn] as Excel.Range).Value;
                    textDestination = (string)(xlWorkSheet_2.Cells[line, startColumn + 1] as Excel.Range).Value;

                    if (textDestination != null && textSource != null)
                    {
                        textSource = textSource.Replace('=', ' ');
                        textDestination = textDestination.Replace('=', ' ');

                        if (textDestination.Contains(item.eplanName) || textSource.Contains(item.eplanName))
                        {
                            //sirve si tenemos un gerate de importancia(de importar en TIA) tanto si está en ziel como en destination


                            xlWorksheet.Cells[newLine, 1] = name;
                            xlWorksheet.Cells[newLine, 5] = textSource;
                            xlWorksheet.Cells[newLine, 6] = textDestination;
                            

                            if (textSource.Contains(item.eplanName)) //is a Exit from the PLC. Take the adresse
                            {
                                position = textSource.IndexOf(":"); //look for that character
                                positionSen = textDestination.IndexOf("+");
                                positionSenFin = textDestination.IndexOf(":");
                                senactadd = textDestination.Substring(positionSen, 10);
                                hdwadd = textSource.Substring(position+1); //maybe we have to put lenght-1. 
                                foreach (SW_HW sftwhw in item.adresseSW_HW) //look for the hardware Adress.
                                {
                                    if (sftwhw.hwAd == hdwadd)
                                    {
                                        sfadd = sftwhw.swAd;
                                        xlWorksheet.Cells[newLine, 2] = sfadd;
                                        xlWorksheet.Cells[newLine, 3] = hdwadd;
                                        xlWorksheet.Cells[newLine, 4] = senactadd; 
                                    }
                                }
                            }

                            else  //thats an entrance in the PLC 
                            {
                                position = textDestination.IndexOf(":"); //look for that character
                                positionSen = textSource.IndexOf("+");
                                positionSenFin = textSource.IndexOf(":");
                                hdwadd = textDestination.Substring(position+1); //maybe we have to put lenght-1. 
                                senactadd = textSource.Substring(positionSen, 10); //maybe i have to change it 
                                foreach (SW_HW sftwhw in item.adresseSW_HW) //look for the hardware Adress.
                                {
                                    if (sftwhw.hwAd == hdwadd)
                                    {
                                        sfadd = sftwhw.swAd;
                                        xlWorksheet.Cells[newLine, 2] = sfadd;
                                        xlWorksheet.Cells[newLine, 3] = hdwadd;
                                        xlWorksheet.Cells[newLine, 4] = senactadd;
                                    }
                                }
                            }
                            newLine++; //the next we write in the next position.
                            
                        }
                    }

                    else if (textDestination == null && textSource == null)
                    {
                        finished = true;
                        line = 2;
                    }


                    line++; //read the next line in the 2nd excel. 

                } while (!finished);

                finished = false;
            }

            //at the end

            xlWorksheet.Columns.AutoFit();
        }



        //Compare the data of the EPLAN and the TIA Export. 
       

        public void CompareEPLANTIA(Excel.Worksheet xlWorksheet)
        {
           
            string numIO=null;//IO number from the TIA export
            string desIO = null; //Description IO from the TIA export
            string numIO_1 = null;//IO number from the EPLAN export
            string desIO_1 = null; //Description IO from the EPLAN export
            int line=1;
            bool finished = false;
            Excel.Range currentFind=null;
            Excel.Range range1;
            conexionEPLAN data = new conexionEPLAN("null", "null", "null", "null", "null"); //create a new linked list. 


            //to use the new Excel 4. 

            do
            {

                numIO = (string)(xlWorkSheet_4.Cells[line, 1] as Excel.Range).Value;
                desIO = (string)(xlWorkSheet_4.Cells[line, 2] as Excel.Range).Value;


                //take the second and the forth column from the excel. 

                range1 = xlWorksheet.Columns["B:B"] as Excel.Range;

                //search for that data. In currentFind will be true/false 

                try
                {
                    currentFind = range1.Find(numIO);//look for numIO on the excel 3. 
                   

                    if (currentFind.Find(numIO) != null) //at least one exist. 
                    {
                        numIO_1 = (string)(xlWorksheet.Cells[currentFind.Row,2] as Excel.Range).Value; //will be the same value as numIO
                        desIO_1 = (string)(xlWorksheet.Cells[currentFind.Row,4] as Excel.Range).Value; //take what name have 


                        if (desIO == null && numIO == null) //when all the data is null in th excel export from TIA
                        {
                            finished = true;
                        }

                        if (desIO!=desIO_1) //Data of the Adress and the name are not correct. 
                        {
                            data = new conexionEPLAN(numIO,desIO,numIO_1,desIO_1, "The adress has not the correct name."); //create a new variable to refer to that in the memory (we work with List)
                            dataNotSimilar.Add(data);
                           

                            xlWorkSheet_4.Cells[line,4].Interior.ColorIndex= 45;
                            xlWorkSheet_4.Cells[line,4] = "The adress has not the correct name."; 
                        }
                    }

                    
                }
                catch
                {

                    if (desIO == null && numIO == null) //when all the data is null in th excel export from TIA
                    {
                        finished = true;
                    }

                    data = new conexionEPLAN(numIO, desIO, "", "", "The adress doesn´t appear in EPLAN"); //create a new variable to refer to that in the memory (we work with List)
                    dataNotSimilar.Add(data);
                    xlWorkSheet_4.Cells[line, 4].Interior.ColorIndex = 3;
                    xlWorkSheet_4.Cells[line, 4] = "The adress doesn´t appear in EPLAN";
                }


                line++;
            } while (!finished);

            xlWorkSheet_4.Columns.AutoFit(); 
        }

        //Show the data that its not the similar in both. 

        public void ShowEnd()
        {
            Form2 form2 = new Form2();
            
            form2.ShowDialog();
            
        }

        //Save and close Excel
        public void SaveCloseExcel(string savePath,Excel.Application xlApp, Excel.Workbook xlWorkBook1, Excel.Worksheet xlSheet, object misValue)
        {
           Excel.Application xlAPP= xlApp;
           Excel.Workbook xlWorkbook1= xlWorkBook1;
           Excel.Worksheet xlWorksheet1= xlSheet;

           string pathString = System.IO.Path.Combine(savePath, "SubFolder");
           //string filename = System.IO.Path.GetFileName(savePath);
            

            if (Directory.Exists(pathString))
            {

                System.IO.Directory.CreateDirectory(pathString);

            }
             

            if (savePath.Contains(".xlsx") && !Directory.Exists(savePath)) 
            {
                savePath = savePath.Replace(".xlsx", " ");
                savePath = savePath + "_neue.xlsx";
            }
            if (savePath.Contains(".xlsm") && !Directory.Exists(savePath))
            {
                savePath = savePath.Replace(".xlsm", " ");
                savePath = savePath + "_original.xlsm";
            }


            xlWorkbook1.SaveCopyAs(savePath); //we select in each iteration what excel data want to save

            //Close the book and the APP. 

            xlWorkbook1.Close(true);
            xlAPP.Quit();

            //Kill the processes

            Marshal.ReleaseComObject(xlWorkbook1);
            Marshal.ReleaseComObject(xlWorksheet1);
            Marshal.ReleaseComObject(xlAPP);
            KillSpecificExcelFileProcess(savePath);
        }

        //-- Functions for TIA Portal 

        //Start instance TIA 

        public void StartTIA()

        {
            
                
                    MyTiaPortal = new TiaPortal(TiaPortalMode.WithoutUserInterface);
                    txt_Status.Text = "TIA Portal started with user interface"; // visualizacion en la interfaz de que se abre con/sin interfaz
           

        }


        //Create new Project 

        public void CreateProject()
        {

            ProjectComposition projectComposition = MyTiaPortal.Projects;

            //Create a new folder with the Project

            DirectoryInfo targetDirectory = new DirectoryInfo(txt_Path5.Text); //take the directory that the user had introduced.

            // Create a project with name Myproject

            Project project = projectComposition.Create(targetDirectory, projectName); //take the name from the excel. 
        }

        private void AddDisp (string serNum, string nameDisp, string CPUPPAL, Project project) //nameDisp: for each eplanName ; CPUPPAL: exported from EPLAN
        {
            DeviceComposition currentDevices = project.Devices;
            //int numberComp = 0; //the first device will be the number 0. LATER WE HAVE TO LOOK IF WE HAVE MORE DISPLAYS IN THE PROJECT
            int pos = 0; //position in the rack
            Device newDevice;
            //HardwareObject newCPU;
            
            
            //IoSystem IO_Moduls= null; 
            //IoConnector IO_Connector;
            //IoControllerComposition ioControllers;


            if (nameDisp == CPUPPAL) 

            {
                 newDevice = currentDevices.CreateWithItem(serNum, nameDisp, "New Device"); //Create a new device. (CPU in the net topology)               
                 //newRack.Name = nameDisp.Substring(0, nameDisp.IndexOf("-")); 
                 //newCPU= newDevice.Items[numberComp]; //create inside the first Rack a PLC
                 //numberComp++; //increment the position of the next device
                 //SoftwareContainer plcSft = newPLC.Soft; //esto es para programar cada una de las CPu.


            }

             // not the CPU, add Modules to the CPU. 
            else
            {
                if (currentDevices != null) //if we have at least 1 CPU. 

                {

                    //newDevice = currentDevices.CreateWithItem(serNum, nameDisp, "New Device"); //Create a new device. (CPU)
                    //newRack = newDevice.Items[numberComp]; //obtein the current Rack
                    
                    foreach(Device device in currentDevices)
                    {   
                        
                            
                                if (device.Name == CPUPPAL) //in CPUPPAL we have the information of the CPU PPAL for each Modul from the excel(IO or similar) 
                                {


                                    if (device.CanPlugNew(serNum, nameDisp, pos))
                                    {
                                        DeviceItem newIO = device.PlugNew(serNum, nameDisp,pos);
                                    }
                               
                                    pos++; //increment the position of the next device.
                                }
                 
                    }


                }

            }
                

        }

        //Not yet implemented. 
        private void VerifyHardware (Device device) //Verify if for each CPU all the devices were introduced.  
        {

            DeviceItemComposition deviceItemComposition = device.DeviceItems;
            if (deviceItemComposition != null)

            {

                foreach (DeviceItem deviceItem in deviceItemComposition)
                {
                    // add co
                }

            }
            


        }

        //Kill the TIA processes. 
        
        public void DisposeTIA()
        {
            //Creamos linked list con variables de tipo Process. Siempre vamos a crear liked list cuando tengamos numero de datos indefinido. 

            IList<TiaPortalProcess> tiaProcessList = TiaPortal.GetProcesses(); //Obtenemos cuantos procesos de TIA hay abiertos en un momento

            if (tiaProcessList.Count > 0) //si hay mas de un proceso abierto en TIA. 
            {
                //Show YES/NO warning that says that ALL the TIA Portal 15.1 instances will be closed

                DialogResult dialogResult = MessageBox.Show(closeAllInstances, closeAllInstancesWarning, MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                if (dialogResult == DialogResult.Yes)
                {
                    //Update status message
                    txt_Status.Text = "TIA Portal closing";
                    MyTiaPortal.Dispose();
                }
                else
                {
                    txt_Status.Text = "TIA Portal are not going to close";

                }
            }

        }
        
        //Search Project (Implement if we have a created Project, add direct the PLC or the element) 

      
        //Search Zip Project


        public void OpenProject(string ProjectPath)
        {
            try
            {
                MyProject = MyTiaPortal.Projects.Open(new FileInfo(ProjectPath));
               
            }
            catch (Exception ex)
            {
                txt_Status.Text = "Error while opening project" + ex.Message;
             
            }
           
        }

        public void SaveProject()
        {
            MyProject.Save();
            txt_Status.Text = "Project saved";
        }


        private void CloseTIAInstance()
        {
        
            //Try to close a project if it is opened
            try
            {
                //Close project if it is opened
                MyProject.Close(); 
            }
            catch
            {

            }
            //Desconectamos del TIA Portal. 
            MyTiaPortal.Dispose();
            //Cerramos todas las instancias de TIA Portal abiertas
            foreach (var process in Process.GetProcessesByName("Siemens.Automation.Portal"))
            {
                process.Kill();
            }
            //Indicamos que el TIA se ha cerrado. 

        }

        //------------Assembly to connect with TIA. 

        private static Assembly MyResolver(object sender, ResolveEventArgs args)
        {
            int index = args.Name.IndexOf(',');
            if (index == -1)
            {
                return null;
            }
            string name = args.Name.Substring(0, index);

            RegistryKey filePathReg = Registry.LocalMachine.OpenSubKey(
                "SOFTWARE\\Siemens\\Automation\\Openness\\15.1\\PublicAPI\\15.1.0.0");

            if (filePathReg == null)
                return null;

            object oRegKeyValue = filePathReg.GetValue(name);
            if (oRegKeyValue != null)
            {
                string filePath = oRegKeyValue.ToString();

                string path = filePath;
                string fullPath = Path.GetFullPath(path);
                if (File.Exists(fullPath))
                {
                    return Assembly.LoadFrom(fullPath);
                }
            }

            return null;
        }


        //English selected
        private void btn_EN_Click(object sender, EventArgs e)
        {
            //If previous language was german
            if (languageSelected == "DE")
            {
                //Set language as english
                languageSelected = "EN";
                //Change texts to english
                languageEN();
            }
        }

        //German selected
        private void btn_DE_Click(object sender, EventArgs e)
        {
            //If previous language was english
            if (languageSelected == "EN")
            {
                //Set language as german
                languageSelected = "DE";
                //Change texts to german
                languageDE();
            }
        }

        //Change texts to english
        private void languageEN()
        {
            //Change background color of the language buttons 
            btn_EN.BackColor = System.Drawing.Color.FromArgb(164, 13, 48);
            btn_DE.BackColor = System.Drawing.Color.FromArgb(215, 25, 70);

            //Load texts in english

            //Strings for messages

            errorMessage = errorMessageEN;
            executing = executingEN;
            programmFinished = programmFinishedEN;
            programmStopped = programmStoppedEN;
            nothingFound = nothingFoundEN;
            invalidPath = invalidPathEN;

            //Strings for buttons

            txtStatusLabel.Text = statusEN;
            txtSelect1.Text = txtPathField1EN;
            txtSelect2.Text = txtPathField2EN;
            txtSelect3.Text = txtPathField3EN;
            txtSelect4.Text= txtPathField4EN;
            btn_Path1.Text = browseEN;
            btn_Path2.Text = browseEN;
            btn_Path3.Text = browseEN;
            button1.Text = browseEN; 
            


            //Change the text in txt_Status to english
            string text = txt_Status.Text;
            if (text != string.Empty)
            {
                if (text == executingDE)
                {
                    text = executingEN;
                }
                else if(text == programmFinishedDE)
                {
                    text = programmFinishedEN;
                }
                else if (text == programmStoppedDE)
                {
                    text = programmStoppedEN;
                }
                else if (text == nothingFoundDE)
                {
                    text = nothingFoundEN;
                }
            }
            txt_Status.Text = text;



            if (!String.IsNullOrEmpty(label3.Text))
            {
                label3.Text = invalidPath;
            }
            if (!String.IsNullOrEmpty(label5.Text))
            {
                label5.Text = invalidPath;
            }
        }
        //Change texts to german
        private void languageDE()
        {
            //Change background color of the language buttons 
            btn_DE.BackColor = System.Drawing.Color.FromArgb(164, 13, 48);
            btn_EN.BackColor = System.Drawing.Color.FromArgb(215, 25, 70);

            //Load texts in german

            //Strings for messages
            errorMessage = errorMessageDE;

            executing = executingDE;
            programmFinished = programmFinishedDE;

            programmStopped=programmStoppedDE;
            nothingFound = nothingFoundDE;
            longNamesTextBeggin =longNamesTextBegginDE;
            longNamesTextEnd=longNamesTextEndDE;
            
            invalidPath = invalidPathDE;

            //Strings for buttons

           
            txtSelect1.Text = txtPathField1DE;
            txtSelect2.Text = txtPathField2DE;
            txtSelect3.Text = txtPathField3DE;
            txtSelect4.Text = txtPathField4DE;
            btn_Path1.Text = browseDE;
            btn_Path2.Text = browseDE;
            btn_Path3.Text = browseDE;
            button1.Text = browseDE;


            //Change the text in txt_Status to german
            string text = txt_Status.Text;
            if (!String.IsNullOrEmpty(text))
            {
                if (text == executingEN)
                {
                    text = executingDE;
                }
                else if (text == programmFinishedEN)
                {
                    text = programmFinishedDE;
                }
                else if (text == programmStoppedEN)
                {
                    text = programmStoppedDE;
                }
                else if (text== nothingFoundEN)
                {
                    text = nothingFoundDE;
                }
            }
            txt_Status.Text = text;



            if (!String.IsNullOrEmpty(label3.Text))
            {
                label3.Text = invalidPath;
            }
            if (!String.IsNullOrEmpty(label5.Text))
            {
                label5.Text = invalidPath;
            }
        }

        private void KillSpecificExcelFileProcess(string excelFileName)
        {
            var processes = from p in Process.GetProcessesByName("EXCEL")
                            select p;

            foreach (var process in processes)
            {
                if (process.MainWindowTitle == "Microsoft Excel - " + excelFileName)
                    process.Kill();
            }
        }

        [DllImport("user32.dll", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();

        [DllImport("user32.dll", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr one, int two, int three, int four);
        //Close visualization
        private void button_Exit_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }
        //Maximize visualization
        private void button_Maximize_Click(object sender, EventArgs e)
        {
            if (this.WindowState == (FormWindowState)2) { WindowState = FormWindowState.Normal; } //Set visualisation to normal size
            else { WindowState = FormWindowState.Maximized; } //Maximize visualisation
        }
        //Minimaze visualization
        private void button_Minimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        //Move visualization
        private void panel3_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(Handle, 0x112, 0xf012, 0);
        }


        private void btn_Path_Click(object sender, EventArgs e)
        {
            //Open a File Dialog
            var dialog = new VistaOpenFileDialog();
            //Set filter to show only Excel files
            dialog.Filter = @"*.xlsx|*.xlsx|(*.xls)|*.xls|(*.xlsm)|*.xlsm |(*.xlsb)|*.xlsb";
            //Show Dialog
            dialog.ShowDialog();
            //Get the complete path of the Excel file to be opened
            string excelPath_1 = dialog.FileName;
            //Show path selected in the upper textbox
            txt_Path1.Text = excelPath_1;
        }

        private void btn_Path2_Click(object sender, EventArgs e)
        {
            VistaFolderBrowserDialog dialog = new VistaFolderBrowserDialog();
            dialog.ShowDialog();
            string excelPath_2 = dialog.SelectedPath.ToString();
            //Show selected path
            txt_Path2.Text = excelPath_2;
        }

        
        private void btn_Path3_Click(object sender, EventArgs e)
        {
            //Open a File Dialog
            var dialog = new VistaOpenFileDialog();
            //Set filter to show only Excel files
            dialog.Filter = @"*.xlsx|*.xlsx|(*.xls)|*.xls|(*.xlsm)|*.xlsm|(*.xlsb)|*.xlsb";
            //Show Dialog
            dialog.ShowDialog();
            //Get the complete path of the Excel file to be opened
            string excelPath_3 = dialog.FileName;
            //Show path selected in the upper textbox
            txt_Path3.Text = excelPath_3;
        }

        private void btn_Path5_Click(object sender, EventArgs e)
        {
            VistaFolderBrowserDialog dialog = new VistaFolderBrowserDialog();
            dialog.ShowDialog();
            string txtPath_5 = dialog.SelectedPath.ToString();
            //Show selected path
            txt_Path5.Text = txtPath_5;
        }


        private void txt_Path_TextChanged(object sender, EventArgs e)
        {
            //If txt_Path1 contains a valid path
            if (!String.IsNullOrEmpty(txt_Path1.Text) && System.IO.File.Exists(txt_Path1.Text))
            {
                label3.Text = "";
                if((!String.IsNullOrEmpty(txt_Path2.Text) && System.IO.Directory.Exists(txt_Path2.Text)))
                {
                    //Enable button "Start"
                    btn_Start.Enabled = true;
                }
                
            }
            else if(!String.IsNullOrEmpty(txt_Path1.Text) && !System.IO.File.Exists(txt_Path1.Text))
            {
                label3.Text = invalidPath;
                //Disable button "Start"
                btn_Start.Enabled = false;
            }
            else
            {
                label3.Text = "";
                //Disable button "Start"
                btn_Start.Enabled = false;
            }
        }

        //aqui se introduce donde se quiere exportar el formato de los datos (Path)
        private void txt_Path2_TextChanged(object sender, EventArgs e)
        {
            //If txt_Path2 and txt_Path1 contains a valid path
            if (!String.IsNullOrEmpty(txt_Path2.Text) && System.IO.Directory.Exists(txt_Path2.Text))
            {
                label5.Text = "";
                if (!String.IsNullOrEmpty(txt_Path1.Text) && System.IO.File.Exists(txt_Path1.Text)&& !String.IsNullOrEmpty(txt_Path3.Text) && System.IO.File.Exists(txt_Path3.Text))
                {
                    //Enable button "Start"
                    btn_Start.Enabled = true;
                }

            }
            else if (!String.IsNullOrEmpty(txt_Path2.Text) && !System.IO.Directory.Exists(txt_Path2.Text))
            {
                label5.Text = invalidPath;
                //Disable button "Start"
                btn_Start.Enabled = false;
            }
            else if (!String.IsNullOrEmpty(txt_Path3.Text) && !System.IO.Directory.Exists(txt_Path3.Text))
            {
                label4.Text = invalidPath;
                //Disable button "Start"
                btn_Start.Enabled = false;
            }
            else
            {
                label5.Text = "";
                //Disable button "Start"
                btn_Start.Enabled = false;
            }
        }
      
        //aqui se introduce el backup completo de los programas de un modelo (M2) 
        private void txt_Path3_TextChanged(object sender, EventArgs e)
        {

        }


        private void button1_Click(object sender, EventArgs e)
        {
            //Open a File Dialog
            var dialog = new VistaOpenFileDialog();
            //Set filter to show only Excel files
            dialog.Filter = @"*.xlsx|*.xlsx|(*.xls)|*.xls|(*.xlsm)|*.xlsm|(*.xlsb)|*.xlsb";
            //Show Dialog
            dialog.ShowDialog();
            //Get the complete path of the Excel file to be opened
            string excelPath_4 = dialog.FileName;
            //Show path selected in the upper textbox
            txt_Path4.Text = excelPath_4;
        }

        
    }
}
