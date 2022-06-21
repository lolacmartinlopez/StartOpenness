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
    public partial class Form2 : Form
    {

        public Form2()
        {
            InitializeComponent();
            AppDomain CurrentDomain = AppDomain.CurrentDomain;
            CurrentDomain.AssemblyResolve += new ResolveEventHandler(MyResolver);
        }

        //Variables for the entire Form 
        
        //Language selected (English by default)
        string languageSelected = "EN";


        //English variables
        string errorMessageEN = "Error: ";

        string executingEN = "Programm is in execution...";

        string programmFinishedEN = "The programm has finished";
        string programmStoppedEN = "The programm has been stopped due to an error: there are variables with names longer than 15" +
            " characters. Please fix it and try again.";

        string nothingFoundEN = "No information was found in the Excel. Please check it and try again.";

        string longNamesTextBegginEN = "The following variables have a name longer than 15 characters: \n \n";
        string longNamesTextEndEN = "\nPlease fix it and try again.";

        string txtPathField1EN = "Select path where the Excel file is located";
        string txtPathField2EN = "Select path where the .lc files will be saved";
        string txtCheckBoxEN = "Use Excels folder";

        string browseEN = "Browse";

        string statusEN = "Status";
        string invalidPathEN = "Invalid path";

        //new English variables 

        string closeAllInstances = "All TIA Portal instances will be closed. Do you want to continue?";
        string closeAllInstancesWarning = "Close all TIA Portal instances";

        //German variables
        string errorMessageDE = "Fehler: ";

        string executingDE = "Programm wird ausgeführt...";

        string programmFinishedDE = "Das Programm ist beendet";

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

        List<conexionEPLAN> dataNotSimilar= new List<conexionEPLAN>();


        //Load form
        private void Form2_Load(object sender, EventArgs e)
        {

            //Set visualisation to normal size
            WindowState = FormWindowState.Normal;
            //Load texts in english by default
            languageEN();
            

        }

        public void DisplayInfo(List <conexionEPLAN> dataNotSimilar)
        {
            listBox1.Items.Clear();
            listBox1.DataSource = null;

            foreach (conexionEPLAN conexion in dataNotSimilar)
            {
                if (conexion == null)
                {

                    listBox1.Items.Add(conexion);
                    
                }
                //ListViewItem itemAdd = new ListViewItem();
                //itemAdd.Checked = true;
                //itemAdd = listBox1.Items.Add(conexion.pinIO + conexion.modIO);
                //itemAdd.SubItems.Add(conexion.pinIO_1 + conexion.modIO_1);
                //itemAdd.SubItems[1].Text = conexion.pinIO_1 + conexion.modIO_1; 
                //listView1.Items.AddRange(new ListViewItem[] { itemAdd });


            }
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
        //Load text in English (just add information)

    }
    //Change texts to german
    private void languageDE()
    {
        //Change background color of the language buttons 
        btn_DE.BackColor = System.Drawing.Color.FromArgb(164, 13, 48);
        btn_EN.BackColor = System.Drawing.Color.FromArgb(215, 25, 70);
        //Load text in German 
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


    }
}


