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
            

            foreach (conexionEPLAN conexion in dataNotSimilar)
            {
                if (conexion == null)
                {

                    
                    
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

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}


