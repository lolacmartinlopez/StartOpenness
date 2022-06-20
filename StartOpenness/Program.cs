using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EPLAN_TIA
{
    
    static class Program
    {
        /// <summary>
        /// Der Haupteinstiegspunkt für die Anwendung.
        /// </summary>
        [STAThread]
        static void Main()
        {
            //variables used by the two forms 

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }

    //Global Data (accesible from all the forms)
    public class conexionEPLAN //aqui podemos direccionar sabiendo ya los datos de cada SPS (es decir, como se llama cada uno) Esto será para importarlo en TIA

    {
        //public string codPLC { get; set; } //(FUTURA IMPLEMENTACION)aqui guardamos solo el codigo identificativo del PLC al que pertenece 
        public string modIO { get; set; } //aqui guardamos a cuales de los perifericos pertenece
        public string pinIO { get; set; }  //aqui guardamos el dato de que PIN dentro del modIO iría conectado.
        public string modIO_1 { get; set; }  //aqui guardamos el dato de que PIN dentro del modIO iría conectado.
        public string pinIO_1 { get; set; }  //aqui guardamos el dato de que PIN dentro del modIO iría conectado.


        public conexionEPLAN(/*string codplc*/  string pinio, string modio, string pinio_1, string modio_1)
        {

            //codPLC = codplc; 
            modIO = modio;
            pinIO = pinio;
            modIO_1 = modio_1;
            pinIO_1 = pinio_1;
        }
    }


}
