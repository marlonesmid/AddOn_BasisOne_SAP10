using System;

using System.Windows.Forms;

namespace BasisOne
{
    class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {

            SAPMain SAPMain =new SAPMain();
            SAPMain.Init();

            Application.Run();
        }        
    }
}
