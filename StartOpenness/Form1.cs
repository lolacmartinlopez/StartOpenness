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
using Microsoft.Office.Interop.Excel; 

namespace StartOpenness
{
    public partial class Form1 : Form
    {
        //Excel variables
        Excel._Application xlApp; //Asi llamamos al tipo de formato que me viene con el excel.
        Excel._Application xlApp_2; //Asi llamamos al tipo de formato que me viene con el excel.
        object misValue = System.Reflection.Missing.Value;
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        Excel.Workbook xlWorkBook_2;
        Excel.Worksheet xlWorkSheet_2;


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

        //Datos relevantes 

        //List<string> usedNames = new List<string>();
       
        public Form1()
        {
            InitializeComponent();
            AppDomain CurrentDomain = AppDomain.CurrentDomain;
        }


        public class InformationSPS
        {
            string serialNum;
            string eplanBemerk;
            string IP; //quizas luego tengo que ponerlo modo INT 
            string cpuPPAL; //a que CPU esta conectada
            string startAdresse; //start adresse de la tarjeta
            Dictionary<string,string> adresseSW_HW; //guardaremos los datos de cada una de las conexiones físicas (E0.1... usw) con el dato que tiene en hardware (en el eplan) 
            
        }

        //Load form
        private void Form1_Load(object sender, EventArgs e)
        {
            //Set visualisation to normal size
            WindowState = FormWindowState.Normal;
            //Load texts in english by default
            languageEN();

        }
       
        //Start the process
        private void btn_Start_Click(object sender, EventArgs e)
        {
            //Excel variables parte 1 del proceso

            //xlApp = new Excel.Application();
            //xlWorkBook = xlApp.Workbooks.Open(txt_Path3.Text); //se abre primero el Excel con los programas
            //xlWorkSheet = xlWorkBook.Worksheets.get_Item(1); //abrimos solo la primera hoja del excel



            ////aqui se llama a la funcion leer, y le pasamos la lista para que guarde los valores

            //usedNames= readProgramm(); 

           //Excel puntos parte 2 del proceso

            xlApp_2 = new Excel.Application(); 
            xlWorkBook_2 = xlApp_2.Workbooks.Open(txt_Path1.Text); //se abre primero el Excel con los programas
            xlWorkSheet_2 = xlWorkBook_2.Worksheets.get_Item(1); //abrimos solo la primera hoja del excel

            bool lengthOK;
            int startColumn=2;

            //Show message
            txt_Status.Text = executing;

            //Check if the length of the variable names is correct
            lengthOK=checkLength();
            //If the lengths are correct 
            if(lengthOK)
            {
                //Look for "J1" or "RX" to see which process must be done
               
                    //Read all the joints and create the .lc file
                    readWriteJ1J4J6(startColumn);
                    //Show message
                    txt_Status.Text = programmFinished;
                
              
            }
            //If the lengths are not correct
            else
            {
                //Show error message
                txt_Status.Text = programmStopped;
            }

            //Close Excel file 1
            //xlWorkBook.Close(false, misValue, misValue);
            //Close Excel app and release all
            //xlApp.Quit();
            //Marshal.ReleaseComObject(xlWorkSheet);
            //Marshal.ReleaseComObject(xlWorkBook);
            //Marshal.ReleaseComObject(xlApp);

            //Close Excel file 2
            xlWorkBook_2.Close(false, misValue, misValue);
            //Close Excel app and release all
            xlApp_2.Quit();
            Marshal.ReleaseComObject(xlWorkSheet_2);
            Marshal.ReleaseComObject(xlWorkBook_2);
            Marshal.ReleaseComObject(xlApp_2);


        }
        //Check the lengths of the variable names and create a list of the variables whose name is too long
        private bool checkLength()
        {
            bool lengthOK=true;
            bool finished = false;
            string type;
            string name;
            string longNamesList="";
            int line = 2;

            do
            {
                //Read the type of data
                type = (string)(xlWorkSheet_2.Cells[line, 1] as Excel.Range).Value;
                //If it was an empty cell, the check process is over
                if (String.IsNullOrEmpty(type))
                {
                    finished = true;
                }
                //If it had this text, it was a variable
                else if (type == "PmViaLocationOperation")
                {
                    //Save the variable name
                    name = (string)(xlWorkSheet_2.Cells[line, 1] as Excel.Range).Value;
                    //Check if the length is more than 15 characters
                    if (name.Length > 200)
                    {
                        //Deactivate the variable which indicates that the lengths were correct
                        lengthOK=false;
                        //Add the variable name to the list of variables whose name is too long
                        longNamesList = longNamesList + name + "\n";
                    }
                }
                //Increase the line of the Excel to be read
                line++;

            } while (!finished);
            //If there was some incorrect length
            if (!lengthOK)
            {
                //Show error message with all the names of the variables whose name is too long
                MessageBox.Show(longNamesTextBeggin + longNamesList + longNamesTextEnd);
            }
            return lengthOK;
        }

        //Read all the programs of one model of a Robot
        //private List<string>  readProgramm()
        //{
        //    int iCurrentLine=1; //inicalemnte es la primera linea 
        //    string currentLine; //la primera linea 
        //    bool finished=false;
        //    string name;
        //    int nameInit;
        //    int nameFin;
        //    int nameLen; //initial size
        //    List <string> nameList= new List<string>(); //guardamos puntos que quiere el usuario. 

        //    do
        //    {
        //        //Read the type of data
        //        currentLine = (string)(xlWorkSheet.Cells[iCurrentLine, 1] as Excel.Range).Value;

        //        //If it was an empty cell, the check process is over
        //        if (String.IsNullOrEmpty(currentLine))
        //        {
        //            finished = true;
        //        }
        //        //If it had this text, it was a variable
        //        else if (currentLine == "LMOVE")
        //        {
        //            //Save the variable name
                     
        //            name = (string)(xlWorkSheet.Cells[iCurrentLine, 1] as Excel.Range).Value;
        //            nameLen = name.Length;

        //            //Calculo tamano 
        //            nameInit = name.IndexOf("#", 1, name.Length); //posicion de la linea que encontramos el simbolo de la posicion
        //            nameFin = name.IndexOf(";", 1, name.Length); 
                    
        //            name= name.Remove(1,nameInit-1); //borramos LMOVE o JMOVE
        //            name = name.Remove(nameFin, name.Length);

        //            nameList.Add(name);
        //        }

        //        else if (currentLine == "JMOVE")
        //        {
        //            //Save the variable name

        //            name = (string)(xlWorkSheet.Cells[iCurrentLine, 1] as Excel.Range).Value;
        //            nameLen = name.Length;

        //            //Calculo tamano 
        //            nameInit = name.IndexOf("#", 1, name.Length); //posicion de la linea que encontramos el simbolo de la posicion
        //            nameFin = name.IndexOf(";", 1, name.Length);

        //            name = name.Remove(1, nameInit - 1); //borramos LMOVE o JMOVE
        //            name = name.Remove(nameFin, name.Length);

        //            nameList.Add(name);


        //        }
        //        //Increase the line of the Excel to be read
        //        iCurrentLine++;

        //    } while (!finished);

        //    return nameList;
        //}

      


        //Read J1, J4 und J6, change the value and create .lc file
        private void readWriteJ1J4J6(int startColumn) //le pasamos la lista de los unicos puntos que interesan. 
        {   
            string name;
            string baseName = "";
            int j1Pos;
            int j2Pos;
            int j3Pos;
            int j4Pos;
            int j5Pos;
            int j6Pos;
            double j1;
            double j2;
            double j3;
            double j4;
            double j5;
            double j6;

            bool finished = false;
            //bool previousBase = false;
            int line = 2; //the first line ist 2 (the 1 ist only names)
            string cellText = "";
            string text = ".JOINTS\n";
            string ceros = "0.000000 0.000000 0.000000";

            //Get the colums where the information of the joints is
            j1Pos = startColumn; //Column where J1 data ist (es la numero 2)
            j2Pos = j1Pos + 1;
            j3Pos = j2Pos + 1;
            j4Pos = j3Pos + 1;
            j5Pos = j4Pos + 1;
            j6Pos = j5Pos + 1;

            //Repeat the process until an empty cell is found
            do
            {
                try
                {   
                    //Read from the Excel the name of the variable and the joints

                    name = (string)(xlWorkSheet_2.Cells[line, 1] as Excel.Range).Value;

                    
                        j1 = (double)(xlWorkSheet_2.Cells[line, j1Pos] as Excel.Range).Value;
                        j2 = (double)(xlWorkSheet_2.Cells[line, j2Pos] as Excel.Range).Value;
                        j3 = (double)(xlWorkSheet_2.Cells[line, j3Pos] as Excel.Range).Value;
                        j4 = (double)(xlWorkSheet_2.Cells[line, j4Pos] as Excel.Range).Value;
                        j5 = (double)(xlWorkSheet_2.Cells[line, j5Pos] as Excel.Range).Value;
                        j6 = (double)(xlWorkSheet_2.Cells[line, j6Pos] as Excel.Range).Value;
                        string addText =  name + " " + (-1) * j1 + " " + j2 + " " + j3 + " " + (-1) * j4 + " " + j5 + " " + (-1) * j6 + " " + ceros;
                        
                        //If the information to add is not repeated

                        if (!text.Contains(addText))
                        {
                            //Add the information of the variable and joints to the .lc
                            text = text + addText + "\n";
                        }
                    


                }
                //If the joints of the variable couldn´t be read->empty cell->name of a base
                catch
                {
                    //Get the name of the base or the name of the point
                    cellText = (string)(xlWorkSheet_2.Cells[line, 1] as Excel.Range).Value;
                    //If there was a previous base
                    if (cellText== ".END")
                    {
                        text = text.Replace(",", ".");
                        text = text + ".END";

                        string savePath;
                        //If the path for the .lc file must be folder where the Excel is
                        if (checkBox1.Checked)
                        {
                            //Get the path of the folder where the Excel is
                            savePath = System.IO.Path.GetDirectoryName(txt_Path1.Text);
                        }
                        else
                        {
                            //Get the path that the user gave
                            savePath = txt_Path2.Text;
                        }

                        //Create .lc file
                        string exportName = baseName + "gespiegelt_JOINTS.lc"; //sollen wir wissen ob es gespiegelt
                       
                        string completePath = System.IO.Path.Combine(savePath, exportName);
                        File.WriteAllText(completePath, text);

                        //Prepare variable for the next .lc file
                        text = ".JOINTS\n";

                    }

                    ////State that there has been a base
                    //previousBase = true;

                    //Save the name of the base
                    baseName = cellText;

                    //If the cell is empty
                    if (String.IsNullOrEmpty(cellText))
                    {
                        //Finish the process
                        finished = true;
                    }
                }

                line++;
            } while (!finished);
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
            txtSelect2.Text = txtPathField2EN;
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
            txtSelect2.Text = txtPathField2DE;
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
            dialog.Filter = @"*.as|*.pg";
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
            string lcPath = dialog.SelectedPath.ToString();
            //Show selected path
            txt_Path2.Text = lcPath;
        }

        //introducimos
        //private void btn_Path3_Click(object sender, EventArgs e)
        //{
        //    //Open a File Dialog
        //    var dialog = new VistaOpenFileDialog();
        //    //Set filter to show only Excel files
        //    dialog.Filter = @"*.xlsx|*.xlsx";
        //    //Show Dialog
        //    dialog.ShowDialog();
        //    //Get the complete path of the Excel file to be opened
        //    string excelPath_3 = dialog.FileName;
        //    //Show path selected in the upper textbox
        //    txt_Path3.Text = excelPath_3;
        //}

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

        //aqui se introduce donde se quiere exportar el formato de los datos
        private void txt_Path2_TextChanged(object sender, EventArgs e)
        {
            //If txt_Path1 contains a valid path
            if (!String.IsNullOrEmpty(txt_Path2.Text) && System.IO.Directory.Exists(txt_Path2.Text))
            {
                label5.Text = "";
                if (!String.IsNullOrEmpty(txt_Path1.Text) && System.IO.File.Exists(txt_Path1.Text))
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
            else
            {
                label5.Text = "";
                //Disable button "Start"
                btn_Start.Enabled = false;
            }
        }
      
        //aqui se introduce el backup completo de los programas de un modelo (M2) 
        //private void txt_Path3_TextChanged(object sender, EventArgs e)
        //{
        //    //If txt_Path3 contains a valid path
        //    if (!String.IsNullOrEmpty(txt_Path3.Text) && System.IO.File.Exists(txt_Path3.Text))
        //    {
        //        label5.Text = "";
        //        if (!String.IsNullOrEmpty(txt_Path3.Text) && System.IO.Directory.Exists(txt_Path1.Text))
        //        {
        //            //Enable button "Start"
        //            btn_Start.Enabled = true;
        //        }

        //    }
        //    else if (!String.IsNullOrEmpty(txt_Path3.Text) && !System.IO.Directory.Exists(txt_Path3.Text))
        //    {
        //        label5.Text = invalidPath;
        //        //Disable button "Start"
        //        btn_Start.Enabled = false;
        //    }
        //    else
        //    {
        //        label5.Text = "";
        //        //Disable button "Start"
        //        btn_Start.Enabled = false;
        //    }
        //}

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

        private void btn_Path3_Click(object sender, EventArgs e)
        {

        }
    }
}
