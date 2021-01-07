using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbobsCOM;
using SAPbouiCOM;
using Funciones;
using System.IO;
using System.Reflection;

namespace Intercompany
{
    public class Intercompany
    {

        SAPbouiCOM.Application sboapp;
        SAPbobsCOM.Company oCompany;

        public Intercompany(Application InterSboapp, SAPbobsCOM.Company _company)
        {
            this.sboapp = InterSboapp;
            this.oCompany = _company;

        }

        public void AddFolderToPurchaseRequestForm(SAPbouiCOM.Form oFormPurchaseRequest)
        {
            SAPbouiCOM.Form _oFormPurchaseRequest;
            SAPbouiCOM.Item _oNewItem;
            SAPbouiCOM.Item _oItem;
            SAPbouiCOM.Folder _oFolderItem;

            _oFormPurchaseRequest = oFormPurchaseRequest;
            _oNewItem = _oFormPurchaseRequest.Items.Add("FolderBO1", SAPbouiCOM.BoFormItemTypes.it_FOLDER);
            _oItem = _oFormPurchaseRequest.Items.Item("112");

            _oNewItem.Top = _oItem.Top;
            _oNewItem.Height = _oItem.Height;
            _oNewItem.Width = _oItem.Width;
            _oNewItem.Left = _oItem.Left + _oItem.Width;

            _oFolderItem = ((SAPbouiCOM.Folder)(_oNewItem.Specific));

            _oFolderItem.Caption = "Intercompañia";

            _oFolderItem.GroupWith("112");

            AddItems(_oFormPurchaseRequest);

            _oFormPurchaseRequest.PaneLevel = 1;


        }

        public void AddFolderToPurchaseOrder(SAPbouiCOM.Form oFormPurchaseRequest)
        {
            SAPbouiCOM.Form _oFormPurchaseRequest;
            SAPbouiCOM.Item _oNewItem;
            SAPbouiCOM.Item _oItem;
            SAPbouiCOM.Folder _oFolderItem;

            _oFormPurchaseRequest = oFormPurchaseRequest;
            _oNewItem = _oFormPurchaseRequest.Items.Add("FolderBO2", SAPbouiCOM.BoFormItemTypes.it_FOLDER);
            _oItem = _oFormPurchaseRequest.Items.Item("138");

            _oNewItem.Top = _oItem.Top;
            _oNewItem.Height = _oItem.Height;
            _oNewItem.Width = _oItem.Width;
            _oNewItem.Left = _oItem.Left +_oItem.Left;

            _oFolderItem = ((SAPbouiCOM.Folder)(_oNewItem.Specific));

            _oFolderItem.Caption = "Intercompañia";

            _oFolderItem.GroupWith("138");

            AddItems(_oFormPurchaseRequest);

            _oFormPurchaseRequest.PaneLevel = 1;


        }

        public void ChangePaneFolder(SAPbouiCOM.Form oFormPurchaseRequest)
        {
            SAPbouiCOM.Form _oFormPurchaseRequest;
            _oFormPurchaseRequest = oFormPurchaseRequest;
            _oFormPurchaseRequest.PaneLevel = 5;
        }

        private void AddItems(SAPbouiCOM.Form oFormPurchaseRequest)
        {
            SAPbouiCOM.Item oCamposPurchaseRequest = null;
            SAPbouiCOM.Form _oFormPurchaseRequest;
            SAPbouiCOM.StaticText oStaticText = null;
            SAPbouiCOM.Item oItem;
            SAPbouiCOM.ComboBox oComboBox = null;

            _oFormPurchaseRequest = oFormPurchaseRequest;

            //*******************************************
            // Se adiciona Label "Integrar Documento"
            //*******************************************

            oItem = _oFormPurchaseRequest.Items.Item("62");

            oCamposPurchaseRequest = _oFormPurchaseRequest.Items.Add("blbID", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oCamposPurchaseRequest.Left = oItem.Left + 10;
            oCamposPurchaseRequest.Width = oItem.Width;
            oCamposPurchaseRequest.Top = oItem.Top;
            oCamposPurchaseRequest.Height = oItem.Height;

            oCamposPurchaseRequest.LinkTo = "ComboBox1";

            oStaticText = ((SAPbouiCOM.StaticText)(oCamposPurchaseRequest.Specific));

            oStaticText.Caption = "Integrar Documento ?";

            oCamposPurchaseRequest.FromPane = 5;
            oCamposPurchaseRequest.ToPane = 5;

            //*******************************************
            // Se adiciona Combo box "Integrar Documento"
            //*******************************************

            oCamposPurchaseRequest = _oFormPurchaseRequest.Items.Add("ComboBox1", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            oCamposPurchaseRequest.Left = oItem.Left + 120;
            oCamposPurchaseRequest.Width = oItem.Width;
            oCamposPurchaseRequest.Top = oItem.Top;
            oCamposPurchaseRequest.Height = oItem.Height;

            oCamposPurchaseRequest.DisplayDesc = true;

            _oFormPurchaseRequest.DataSources.UserDataSources.Add("CombSource", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);

            oComboBox = ((SAPbouiCOM.ComboBox)(oCamposPurchaseRequest.Specific));

            
            oComboBox.DataBind.SetBound(true, "", "CombSource");

            oComboBox.ValidValues.Add("Y", "Si");
            oComboBox.ValidValues.Add("N", "No");

            oComboBox.Select("Y", BoSearchKey.psk_ByValue);

            oCamposPurchaseRequest.FromPane = 5;
            oCamposPurchaseRequest.ToPane = 5;
        }

        public void CreacionTablasyCamposeIntercompany()
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            string[] ValidValuesFields1 = { "O", "Origen", "D", "Destino" };
            DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, "", BoYesNoEnum.tNO, ValidValuesFields1, "OPRQ", "BO_TD", "Tipo Doc");
            string[] ValidValuesFields2 = { "Y", "Si", "N", "No" };
            DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, "", BoYesNoEnum.tNO, ValidValuesFields2, "OPRQ", "BO_I", "Integrar Doc?");
            DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, "", BoYesNoEnum.tNO, ValidValuesFields2, "OPRQ", "BO_M", "Doc. Migrado?");
        }

        public string VersionDll()
        {
            try
            {

                Assembly Assembly = Assembly.LoadFrom("BOIntercompany.dll");
                Version vVersion = Assembly.GetName().Version;

                String VersionDll = vVersion.ToString();

                return VersionDll;
            }
            catch (Exception)
            {

                throw;
            }

        }

    }
}
