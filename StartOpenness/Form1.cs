using Microsoft.Win32;
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
    public partial class Form1 : Form
    {
        public Form1()
        {

            InitializeComponent();
            AppDomain CurrentDomain = AppDomain.CurrentDomain;
            CurrentDomain.AssemblyResolve += new ResolveEventHandler(MyResolver);
        }

        //Excel variables

        //Excel excel = new Excel()

        Excel.Application xlApp= new Excel.Application(); //Asi llamamos al tipo de formato que me viene con el excel.
        Excel.Application xlApp_2= new Excel.Application(); //Asi llamamos al tipo de formato que me viene con el excel.
        Excel.Application xlApp_3= new Excel.Application(); //Aqui vamos a guardar un nuevo excel con solo los datos de verbindung de los gearetes que son de interes en cada dispositivo de E/S

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

        string longNamesTextBegginEN="The following variables have a name longer than 15 characters: \n \n";
        string longNamesTextEndEN="\nPlease fix it and try again.";

        string txtPathField1EN = "Select path where the Excel file is located";
        string txtPathField2EN = "Select path where the .lc files will be saved";
        string txtCheckBoxEN= "Use Excels folder";

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

        string txtPathField1DE = "Wähle den Pfad, in dem sich die Excel Datei befindet";
        string txtPathField2DE = "Wähle den Pfad, in dem die .lc Dateien gespeichert werden sollen";
        string txtCheckBoxDE = "Excels Ordner verwenden";

        string browseDE = "Durchsuchen";

        string statusDE = "Status";
        string invalidPathDE = "Ungültiger Pfad";

        //Lista de objetos 

        List<InformationPLC> informationSPs = new List<InformationPLC>(); //keep all the information of the differents SPS avaible
        List <conexionEPLAN> verbindungenEPLANs = new List<conexionEPLAN>(); //keep only the conexions for each IO port from a SPS (or articles that we can import in TIA)
        
        //creacion excel 

      

        //Ejemplo de como guardar la informacion del primer Excel. 
        public class InformationPLC //cada dato de SPS lo escribimos y leemos de esta clase (de ahi que indiquemos metodo get y set 
        {
            public string serialNum { get; set; }
            public string eplanName { get; set; }
            public string IP { get; set; }                                 //quizas luego tengo que ponerlo modo INT 
            public string cpuPPAL { get; set; }                           //a que CPU esta conectada
            public string startAdresse { get; set; }                     //start adresse de la tarjeta
            public List<SW_HW> adresseSW_HW = new List<SW_HW>(); //guardaremos los datos de cada una de las conexiones físicas (E0.1... usw) con el dato que tiene en hardware (en el eplan) 
            
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

        public class conexionEPLAN //aqui podemos direccionar sabiendo ya los datos de cada SPS (es decir, como se llama cada uno) Esto será para importarlo en TIA

        {
            public string codPLC { get; set; } //aqui guardamos solo el codigo identificativo del PLC al que pertenece 
            public string modIO { get; set; } //aqui guardamos a cuales de los perifericos pertenece
            public string pinIO { get; set; }  //aqui guardamos el dato de que PIN dentro del modIO iría conectado.

            public conexionEPLAN (string codplc, string modio, string pinio)
            {

                codPLC = codplc; 
                modIO = modio;
                pinIO = pinio;
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
        //Ejemplo de como guardar la informacion. 
        //verbindungenEPLAN lineSPS= new verbindungenEPLAN(); aqui le deberiamos de darle los datos que tengamos en las lineas. 

        

        //Load form
        private void Form1_Load(object sender, EventArgs e)
        {
            
            //Set visualisation to normal size
            WindowState = FormWindowState.Normal;
            //Load texts in english by default
            languageEN();

        }
       
        //Start the process
        public void btn_Start_Click(object sender, EventArgs e)
        {
            //Excel variables parte 1 del proceso (obtener datos del SPS)



            xlApp = new Excel.Application(); 
            xlWorkBook = xlApp.Workbooks.Open(txt_Path1.Text); //se abre primero el Excel con los programas
            xlWorkSheet = xlWorkBook.Worksheets.get_Item(1); //abrimos solo la primera hoja del excel


           //Excel variables puntos parte 2 del proceso (obtener datos de Verbindungen), pero debe ser abriendo otra ventana de excel 

            xlApp_2 = new Excel.Application();
            xlWorkBook_2 = xlApp_2.Workbooks.Open(txt_Path3.Text); //se abre primero el Excel con los programas
            xlWorkSheet_2 = xlWorkBook_2.Worksheets.get_Item(1); //abrimos solo la primera hoja del excel

            //Excel para exportar y guardar todos los datos. De primeras estará vacío.  

            xlApp_3 = new Excel.Application();

            //Excel.Workbooks xlWorkBook1_3 = xlApp_3.Workbooks;

            /*Excel.Workbook*/  xlWorkBook_3 = xlApp_3.Workbooks.Add();
            /*Excel.Worksheet*/ xlWorkSheet_3 = xlApp_3.Worksheets.Add();



            //xlApp_3= new Excel.Application();
            //Excel.Workbooks xlWorkBook1_3 = xlApp_3.Workbooks;
            //string pathNew = System.IO.Path.Combine(txt_Path2.Text, "NeueVerbindungen.xlsx"); 
            //Excel.Workbook xlWorkbook_3 = xlApp_3.Workbooks.Add();
            //Excel.Worksheet xlWorkSheet_3 = xlApp_3.Worksheets.Add();

            //Excel.Workbook xlWorkBook_3 = this.xlApp_3.Workbooks.Add();
            //Excel.Worksheet xlWorkSheet_3 = xlWorkBook_3.Worksheets.Add(); 

            //----------PRUEBA 1-------------


            //ExcelLibrary.DataSetHelper.CreateWorkbook(path, DataSet dataset); 

            //string pathNew = System.IO.Path.Combine(txt_Path2.Text, "NeueVerbindungen.xlsx"); 

            int startColumn =1; //in Excel DateiSPS must be the first column. In opposite in the excel Verbinden must be the second column. 
            int startColumn_2 = 9; //in Excel Verbindungen the first value ist for the source and the next for the destination
            
            //Show message

            txt_Status.Text = executing;

            //Excel variable for the new file (Excel Verbindungen 2) 

            
            //read the data from "SPSDatei". 

            getinfoPLC(startColumn);

            txt_Status.Text = "The information of the Excel SPS Datei was correct readed";

            //read the second excel but only with the data from the SPS of the previous data. 

            getinfConexion(startColumn_2,informationSPs);

            txt_Status.Text = "The information of the Excel SPS Verbindungen was correct readed";

            //guardamos los datos importantes en un solo excel. Ya teniamos creado de antes el nuevo libro excel donde exportamos las cosas

            txt_Status.Text = "Exporting information to the new Excel";

            exportExcel();

            //Close Excel 

            xlWorkBook.Close(true);
            xlApp.Quit();
            xlWorkBook_2.Close(true);
            xlApp_2.Quit();
            xlWorkBook_3.Close(true);
            xlApp_3.Quit();

            //Kill the processes

            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlApp);


            Marshal.ReleaseComObject(xlWorkBook_2);
            Marshal.ReleaseComObject(xlWorkSheet_2);
            Marshal.ReleaseComObject(xlApp_2);


            //Marshal.ReleaseComObject(xlWorkBook_3);
            //Marshal.ReleaseComObject(xlWorkSheet_3);
            //Marshal.ReleaseComObject(xlApp_3);


            //the third Excel keep opened.

            MyTiaPortal = new TiaPortal(TiaPortalMode.WithoutUserInterface);

            //Create a new TIA Project

            CreateProject();

            //Add the dispositives. 

            AddDisp(); 



            //creamos la correlación de datos entre el excel2 y las salidas de cada uno de los gerätes del programa. 


            //Show message

            txt_Status.Text = programmFinished;
           
        }
        
        //--Functions for the first part of the APP. (Obtein the data of the EPLAN) 

        public void getinfoPLC(int startColumn) //le pasamos en cada excel que columna debe empezar a buscar. 
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
            int line = 2;       //the first line ist 2 (the 1 ist only names)
            string anzahlText;  //controlar si tenemos mas dispositivos relevantes del SPS o no. 
            int colRead = 15;   //aqui debemos de tener en cuenta por que columna vamos leyendo, ya que el numero de E/S es diferente en los perifericos que tengamos.
            int iList = 1;      //write in the list. 
            //Repeat the process until an empty cell oof number of this geaete is found
            do
            {
                try
                {
                   adresseSW_HW1 = new List<SW_HW>(); //reload a new direction in memory for the next values. 

                    //adresseSW_HW1.Clear(); //clear for the next iteration
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
                        startAdresse = (string)(xlWorkSheet.Cells[line, 13] as Excel.Range).Value;

                        //aqui declarar bucle que por cada linea de cada geraete me guarde en el diccionario <direccion SW, direccion HW>

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
                        informationSPS.eplanName = eplanMark;
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
                    

                }
                //If the column of the SPS-Typ is empty
                catch
                {
                    //Controll if there is more SPS Data

                    anzahlText = (string)(xlWorkSheet.Cells[line, 2] as Excel.Range).Value;

                   if (anzahlText ==null)
                    {
                        finished1=true; //there are no more SPS data. 
                    }

                    else
                    {
                        line++; //read the next line
                    }
                }

                txt_Status.Text = " Reading the data of the SPS";

            } while (!finished1);

        }

      
        public void getinfConexion(int startColumn, List<InformationPLC> informationSPs)
        {
            int line = 2; 
            string textDestination;
            string textSource;
            string name; 

            int newLine = 2;
            bool finished=false;
            //int numPLC = informationSPs.Count;
            //int i = 0; 
            
            xlWorkSheet_3.Activate();
            xlWorkSheet_3.Cells[1, 1] = "PLC/Display";
            xlWorkSheet_3.Cells[1, 2] = "Source";
            xlWorkSheet_3.Cells[1, 3] = "Destination";

            //Excel.Range excelSize= (Excel.Range)xlWorkSheet_2.Columns;
            //int excelSizenum = (int)excelSize.ColumnWidth; 
            int count = 0; 

            foreach (InformationPLC item in informationSPs)
            {   
                do
                {
                   
                    textSource = (string)(xlWorkSheet_2.Cells[line, startColumn] as Excel.Range).Value;
                    textDestination = (string)(xlWorkSheet_2.Cells[line, startColumn + 1] as Excel.Range).Value;

                    if (textDestination != null && textSource != null)
                    {
                        if(textDestination.Contains(item.eplanName) || textSource.Contains(item.eplanName))
                        {
                            //delete the symbol "="
                            name = item.eplanName.Replace('=', ' ');
                            textDestination = textDestination.Replace('=', ' '); //the name have "=" at the first position, it makes an error in Excel
                            textSource = textSource.Replace('=', ' ');

                            //sirve si tenemos un gerate de importancia(de importar en TIA) tanto si está en ziel como en destination

                            xlWorkSheet_3.Cells[newLine, 1] = name; 
                            xlWorkSheet_3.Cells[newLine, 2] = textSource;
                            xlWorkSheet_3.Cells[newLine, 3] = textDestination;
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

                count++; //Prueba
                finished = false;
            }
          

        }


        //Export to excel
        public void exportExcel ()
        {
           


            string path = txt_Path2.Text;
            string nameWork = "SPS_neueVerbindungen.xlsx";
            string exportFilename = @"" + path + @"\" + nameWork;


            SaveCloseExcel(txt_Path1.Text, xlApp, xlWorkBook, misValue,1); //save the SPSDATEI
            SaveCloseExcel(txt_Path3.Text, xlApp_2, xlWorkBook_2, misValue,2); //save the SPS conexions
            SaveCloseExcel(exportFilename, xlApp_3, xlWorkBook_3, misValue,3); //Save the new excel with all the info


        }

        //Save and close Excel
        public void SaveCloseExcel(string savePath,Excel.Application xlApp, Excel.Workbook xlWorkBook1,object misValue, int seq)
        {
            Excel.Application xlAPP= xlApp;
           Excel.Workbook xlWorkbook1= xlWorkBook1;

            string pathString = System.IO.Path.Combine(savePath, "SubFolder");
            string filename = System.IO.Path.GetFileName(savePath);
            

            if (Directory.Exists(pathString))
            {

                System.IO.Directory.CreateDirectory(pathString);

            }

       
            if(seq==3)
            {
                
                xlWorkBook1.SaveAs(savePath,Excel.XlFileFormat.xlWorkbookDefault);
          
            }

            else
            {
                savePath = savePath.Replace(".xlsx", " ");
                savePath = savePath + "(1).xlsx";
                xlWorkbook1.SaveCopyAs(savePath); //we select in each iteration what excel data want to save

            }

          
        }

        //-- Functions for TIA Portal 


        //Create new Project 

        public void CreateProject()
        {
            ProjectComposition projectComposition = MyTiaPortal.Projects;

            //Create a new folder with the Project
            DirectoryInfo targetDirectory = new DirectoryInfo(@"D:\TiaProjects");

            // Create a project with name Myproject
            Project project = projectComposition.Create(targetDirectory, "MyProject");
        }

        public void AddDisp()
        {
            DeviceComposition currentDevices = project.Devices;
            Device newDevice = currentDevices.CreateWithItem("OrderNumber:6ES7 510-1DJ01-0AB0/V2.0", "PLC_1", "New Device"); 

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
            longNamesTextBeggin = longNamesTextBegginEN;
            longNamesTextEnd = longNamesTextEndEN;

            invalidPath = invalidPathEN;

            //Strings for buttons
            txtStatusLabel.Text = statusEN;
            //txtSelect2.Text = txtPathField2EN;
            txtSelect1.Text = txtPathField1EN;
            btn_Path1.Text = browseEN;
            btn_Path2.Text = browseEN;
            checkBox1.Text = txtCheckBoxEN;


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
            txtStatusLabel.Text = statusDE;
            //txtSelect2.Text = txtPathField2DE;
            txtSelect1.Text = txtPathField1DE;
            btn_Path1.Text = browseDE;
            btn_Path2.Text = browseDE;
            checkBox1.Text = txtCheckBoxDE;

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
            dialog.Filter = @"*.xlsx|*.xlsx";
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
            dialog.Filter = @"*.xlsx|*.xlsx";
            //Show Dialog
            dialog.ShowDialog();
            //Get the complete path of the Excel file to be opened
            string excelPath_3 = dialog.FileName;
            //Show path selected in the upper textbox
            txt_Path3.Text = excelPath_3;
        }

        private void txt_Path_TextChanged(object sender, EventArgs e)
        {
            //If txt_Path1 contains a valid path
            if (!String.IsNullOrEmpty(txt_Path1.Text) && System.IO.File.Exists(txt_Path1.Text))
            {
                label3.Text = "";
                if((!String.IsNullOrEmpty(txt_Path2.Text) && System.IO.Directory.Exists(txt_Path2.Text))||checkBox1.Checked)
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
            //If txt_Path1 contains a valid path
            if (!String.IsNullOrEmpty(txt_Path3.Text) && System.IO.Directory.Exists(txt_Path3.Text))
            {
                label4.Text = "";
                if (!String.IsNullOrEmpty(txt_Path1.Text) && System.IO.File.Exists(txt_Path1.Text))
                {
                    //Enable button "Start"
                    btn_Start.Enabled = true;
                }

            }
            else if (!String.IsNullOrEmpty(txt_Path3.Text) && !System.IO.Directory.Exists(txt_Path3.Text))
            {
                label4.Text = invalidPath;
                //Disable button "Start"
                btn_Start.Enabled = false;
            }
            else
            {
                label4.Text = "";
                //Disable button "Start"
                btn_Start.Enabled = false;
            }
        }

        //sirve para chequear si queremos guardarlo como un .excel o bien .lc, 
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            //If the path for the .lc file is the folder where the Excel is
            if (checkBox1.Checked)
            {
                //Disable buttons and labels to select a path
                txt_Path2.Enabled = false;
                btn_Path2.Enabled = false;
                label5.Enabled = false;
                //If there is a valid path to an Excel
                if (!String.IsNullOrEmpty(txt_Path1.Text) && System.IO.File.Exists(txt_Path1.Text))
                {
                    //Enable button to start the process
                    btn_Start.Enabled = true;
                }
            }
            //If the path for the .lc file must be given
            else
            {
                //Enable buttons and labels to select a path
                txt_Path2.Enabled = true;
                btn_Path2.Enabled = true;
                label5.Enabled = true;
                //If there is already a valid path where the .lc should be saved and a valid path to an Excel
                if (!String.IsNullOrEmpty(txt_Path2.Text) && System.IO.Directory.Exists(txt_Path2.Text))
                {
                    //Enable button to start the process
                    btn_Start.Enabled = true;
                }
                else
                {
                    //Disable button to start the process
                    btn_Start.Enabled = false;
                }
            }
        }


    }
}
