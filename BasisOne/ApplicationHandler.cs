using System;
using SAPbouiCOM;

namespace BasisOne
{
    internal class ApplicationHandler
    {
        private Application sboapp;

        public ApplicationHandler(Application sboapp)
        {
            this.sboapp = sboapp;
        }

        internal void app_Handler(BoAppEventTypes EventType)
        {

            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    //Exit Add-On
                    System.Windows.Forms.Application.Exit();
                    borraMenu();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    borraMenu();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                   // borraMenu();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                   // borraMenu();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    borraMenu();
                    break;
                default:
                    borraMenu();
                    break;
            }
         
        }

        private void borraMenu()
        {
            //  borrar el menú
            if (sboapp.Menus.Exists("mnuBasisOne"))
            {
                sboapp.Menus.RemoveEx("mnuBasisOne");
            }
        }
    }
}
