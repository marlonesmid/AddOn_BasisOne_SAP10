using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;
using SAPbouiCOM;
using Funciones;
using System.Windows.Forms;
using System.Xml;
using eBilling.ServicioEmisionFE;
using System.IO;
using System.Xml.Serialization;
using System.ServiceModel;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using CrystalDecisions.ReportSource;
using System.Reflection;


namespace eBilling
{
    public class eBillingBO
    {
        #region Variables Globales

        private SAPbouiCOM.Application sboapp;
        private SAPbobsCOM.Company oCompany;

        string sGetLastRecord = null;
        string sGetSeriesNumberInvoice = null;
        string sGetSeriesNumberCreditMemo = null;
        string sGetSeriesNumberDebitMemo = null;
        string sGetTiposOperacion = null;
        string sGetFormattedSearch = null;
        string _IDCategory;
        int Rsd = 0;


        #endregion

        #region Parametros globales TFHKA

        ServicioEmisionFE.ServiceClient serviceClient;
        ServicioAdjuntosFE.ServiceClient serviceClientAdjuntos;

        //Especifica el puerto
        BasicHttpBinding port = new BasicHttpBinding();

        #endregion

        public eBillingBO(SAPbouiCOM.Application eBpsboapp, SAPbobsCOM.Company _company)
        {
            this.sboapp = eBpsboapp;
            this.oCompany = _company;

        }

        private void ItemsDocuments(SAPbouiCOM.Form oFormInvoices, string __TipoDoc)
        {
            #region Variables y Objetos

            SAPbouiCOM.Item oCampoInvoices = null;
            SAPbouiCOM.Form _oFormInvoices;
            SAPbouiCOM.StaticText oStaticText = null;
            SAPbouiCOM.Item oItem;
            SAPbouiCOM.Item _ItemEnviarCopia;
            SAPbouiCOM.Item _ItemEnviarOriginal;

            SAPbouiCOM.EditText otxtCEB = null;
            SAPbouiCOM.EditText otxtCRWS = null;
            SAPbouiCOM.EditText otxtMRWS = null;
            SAPbouiCOM.EditText otxtRPDF = null;
            SAPbouiCOM.EditText otxtCUFE = null;
            SAPbouiCOM.EditText otxtQR = null;
            SAPbouiCOM.EditText otxtAFV = null;
            SAPbouiCOM.ComboBox otxtTN = null;
            SAPbouiCOM.ComboBox otxtTND = null;
            SAPbouiCOM.ComboBox cboS;
            SAPbouiCOM.ComboBox cboPP;
            SAPbouiCOM.ComboBox cboMP;
            SAPbouiCOM.ComboBox cboTDesc;

            _oFormInvoices = oFormInvoices;

            #endregion

            #region Button reenviar factura a la DIAN

            //*******************************************
            // Se adiciona Label "Boton Re-Enviar Fac.Elec"
            //*******************************************

            oItem = _oFormInvoices.Items.Item("62");
            _ItemEnviarCopia = _oFormInvoices.Items.Item("230");

            _ItemEnviarOriginal = _oFormInvoices.Items.Add("BtnEnvi", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            _ItemEnviarOriginal.Left = _ItemEnviarCopia.Left;
            _ItemEnviarOriginal.Top = _ItemEnviarCopia.Top + 20;
            _ItemEnviarOriginal.Width = _ItemEnviarCopia.Width + 30;
            _ItemEnviarOriginal.Height = _ItemEnviarCopia.Height + 10;

            SAPbouiCOM.Button _BtnEnviar = (SAPbouiCOM.Button)_ItemEnviarOriginal.Specific;
            _BtnEnviar.Caption = "Enviar a DIAN";

            #endregion

            #region Button Reenviar E-mail

            //*******************************************
            // Se adiciona Label "Boton Re-Enviar E-mail"
            //*******************************************

            _ItemEnviarOriginal = _oFormInvoices.Items.Add("BtnSM", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            _ItemEnviarOriginal.Left = _ItemEnviarCopia.Left + 142;
            _ItemEnviarOriginal.Top = _ItemEnviarCopia.Top + 20;
            _ItemEnviarOriginal.Width = 100;
            _ItemEnviarOriginal.Height = _ItemEnviarCopia.Height + 5;

            SAPbouiCOM.Button _BtnSM = (SAPbouiCOM.Button)_ItemEnviarOriginal.Specific;
            _BtnSM.Caption = "Reenviar E-Mail";

            #endregion

            #region Campo Comentarios Factura Electronica

            //*******************************************
            // Se adiciona Label "Comentarios Fac.Elec"
            //*******************************************

            oCampoInvoices = _oFormInvoices.Items.Add("blbEBC", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oCampoInvoices.Left = oItem.Left + 10;
            oCampoInvoices.Width = oItem.Width;
            oCampoInvoices.Top = oItem.Top;
            oCampoInvoices.Height = oItem.Height;

            oCampoInvoices.LinkTo = "txtEBC";

            oStaticText = ((SAPbouiCOM.StaticText)(oCampoInvoices.Specific));

            oStaticText.Caption = "Comentarios Fac.Elec";

            oCampoInvoices.FromPane = 5;
            oCampoInvoices.ToPane = 5;

            //*******************************************
            // Se adiciona Tex Box "Comentarios Fac.Elec"
            //*******************************************

            oCampoInvoices = _oFormInvoices.Items.Add("txtEBC", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oCampoInvoices.Left = oItem.Left + 120;
            oCampoInvoices.Width = oItem.Width + 40;
            oCampoInvoices.Top = oItem.Top;
            oCampoInvoices.Height = oItem.Height;

            otxtCEB = ((SAPbouiCOM.EditText)(oCampoInvoices.Specific));

            if (__TipoDoc == "FacturaDeClientes" || __TipoDoc == "NotaDebitoClientes")
            {
                otxtCEB.DataBind.SetBound(true, "OINV", "U_BO_EBC");
            }
            else if (__TipoDoc == "NotaCreditoClientes")
            {
                otxtCEB.DataBind.SetBound(true, "ORIN", "U_BO_EBC");
            }

            oCampoInvoices.FromPane = 5;
            oCampoInvoices.ToPane = 5;

            #endregion

            #region Enviar por Correo 

            //*******************************************
            // Se adiciona Label "Enviar Correo Electronico"
            //*******************************************

            oCampoInvoices = _oFormInvoices.Items.Add("lblEE", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oCampoInvoices.Left = oItem.Left + 275;
            oCampoInvoices.Width = 110;
            oCampoInvoices.Top = oItem.Top;
            oCampoInvoices.Height = oItem.Height;

            oCampoInvoices.LinkTo = "txtEE";

            oStaticText = ((SAPbouiCOM.StaticText)(oCampoInvoices.Specific));

            oStaticText.Caption = "Enviar por Correo ?";

            oCampoInvoices.FromPane = 5;
            oCampoInvoices.ToPane = 5;

            //*******************************************
            // Se adiciona Tex Box "Enviar Correo Electronico"
            //*******************************************

            oCampoInvoices = _oFormInvoices.Items.Add("txtEE", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            oCampoInvoices.Left = oItem.Left + 380;
            oCampoInvoices.Width = 30;
            oCampoInvoices.Top = oItem.Top;
            oCampoInvoices.Height = oItem.Height;
            oCampoInvoices.Enabled = true;

            oCampoInvoices.DisplayDesc = true;

            _oFormInvoices.DataSources.UserDataSources.Add("cboEE", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);

            cboPP = ((SAPbouiCOM.ComboBox)(oCampoInvoices.Specific));

            if (__TipoDoc == "FacturaDeClientes" || __TipoDoc == "NotaDebitoClientes")
            {
                cboPP.DataBind.SetBound(true, "OINV", "U_BO_EE");
            }
            else if (__TipoDoc == "NotaCreditoClientes")
            {
                cboPP.DataBind.SetBound(true, "ORIN", "U_BO_EE");
            }

            cboPP.ValidValues.Add("Y", "Si");
            cboPP.ValidValues.Add("N", "No");

            cboPP.Select("Y", BoSearchKey.psk_ByValue);

            oCampoInvoices.FromPane = 5;
            oCampoInvoices.ToPane = 5;

            #endregion

            #region Campo codigo respuesta facturacion electronica

            //*******************************************
            // Se adiciona Label "Cod. Resp. Fac. Elec"
            //*******************************************

            oCampoInvoices = _oFormInvoices.Items.Add("lblCRWS", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oCampoInvoices.Left = oItem.Left + 10;
            oCampoInvoices.Width = oItem.Width;
            oCampoInvoices.Top = oItem.Top + 20;
            oCampoInvoices.Height = oItem.Height;

            oCampoInvoices.LinkTo = "txtCRWS";

            oStaticText = ((SAPbouiCOM.StaticText)(oCampoInvoices.Specific));

            oStaticText.Caption = "Cod. Resp. Fac. Elec";

            oCampoInvoices.FromPane = 5;
            oCampoInvoices.ToPane = 5;

            //*******************************************
            // Se adiciona Tex Box "Cod. Resp. Fac. Elec"
            //*******************************************

            oCampoInvoices = _oFormInvoices.Items.Add("txtCRWS", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oCampoInvoices.Left = oItem.Left + 120;
            oCampoInvoices.Width = 40;
            oCampoInvoices.Top = oItem.Top + 20;
            oCampoInvoices.Height = oItem.Height;
            oCampoInvoices.Enabled = false;



            otxtCRWS = ((SAPbouiCOM.EditText)(oCampoInvoices.Specific));

            if (__TipoDoc == "FacturaDeClientes" || __TipoDoc == "NotaDebitoClientes")
            {
                otxtCRWS.DataBind.SetBound(true, "OINV", "U_BO_CRWS");

            }
            else if (__TipoDoc == "NotaCreditoClientes")
            {
                otxtCRWS.DataBind.SetBound(true, "ORIN", "U_BO_CRWS");
            }

            oCampoInvoices.FromPane = 5;
            oCampoInvoices.ToPane = 5;

            #endregion

            #region Campo Mensaje respuesta Facturacion electronica

            //*******************************************
            // Se adiciona Label "Mens. Resp. Fac. Elec"
            //*******************************************

            oCampoInvoices = _oFormInvoices.Items.Add("lblMRWS", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oCampoInvoices.Left = 180;
            oCampoInvoices.Width = oItem.Width;
            oCampoInvoices.Top = oItem.Top + 20;
            oCampoInvoices.Height = oItem.Height;

            oCampoInvoices.LinkTo = "txtMRWS";

            oStaticText = ((SAPbouiCOM.StaticText)(oCampoInvoices.Specific));

            oStaticText.Caption = "Mens. Resp. Fac. Elec";

            oCampoInvoices.FromPane = 5;
            oCampoInvoices.ToPane = 5;

            //*******************************************
            // Se adiciona Tex Box "Mens. Resp. Fac. Elec"
            //*******************************************

            oCampoInvoices = _oFormInvoices.Items.Add("txtMRWS", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oCampoInvoices.Left = 290;
            oCampoInvoices.Width = oItem.Width + 100;
            oCampoInvoices.Top = oItem.Top + 20;
            oCampoInvoices.Height = oItem.Height;
            oCampoInvoices.Enabled = false;

            otxtMRWS = ((SAPbouiCOM.EditText)(oCampoInvoices.Specific));

            if (__TipoDoc == "FacturaDeClientes" || __TipoDoc == "NotaDebitoClientes")
            {
                otxtMRWS.DataBind.SetBound(true, "OINV", "U_BO_MRWS");
            }
            else if (__TipoDoc == "NotaCreditoClientes")
            {
                otxtMRWS.DataBind.SetBound(true, "ORIN", "U_BO_MRWS");
            }

            oCampoInvoices.FromPane = 5;
            oCampoInvoices.ToPane = 5;

            #endregion

            #region Campo PDF enviado 

            //*******************************************
            // Se adiciona Label "PDF Enviado"
            //*******************************************

            oCampoInvoices = _oFormInvoices.Items.Add("lblRPDF", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oCampoInvoices.Left = oItem.Left + 10;
            oCampoInvoices.Width = oItem.Width;
            oCampoInvoices.Top = oItem.Top + 40;
            oCampoInvoices.Height = oItem.Height;

            oCampoInvoices.LinkTo = "txtRPDF";

            oStaticText = ((SAPbouiCOM.StaticText)(oCampoInvoices.Specific));

            oStaticText.Caption = "PDF Enviado";

            oCampoInvoices.FromPane = 5;
            oCampoInvoices.ToPane = 5;

            //*******************************************
            // Se adiciona Tex Box "PDF Enviado"
            //*******************************************

            oCampoInvoices = _oFormInvoices.Items.Add("txtRPDF", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oCampoInvoices.Left = oItem.Left + 120;
            oCampoInvoices.Width = oItem.Width + 100;
            oCampoInvoices.Top = oItem.Top + 40;
            oCampoInvoices.Height = oItem.Height;
            oCampoInvoices.Enabled = false;

            otxtRPDF = ((SAPbouiCOM.EditText)(oCampoInvoices.Specific));


            if (__TipoDoc == "FacturaDeClientes" || __TipoDoc == "NotaDebitoClientes")
            {
                otxtRPDF.DataBind.SetBound(true, "OINV", "U_BO_RPDF");
            }
            else if (__TipoDoc == "NotaCreditoClientes")
            {
                otxtRPDF.DataBind.SetBound(true, "ORIN", "U_BO_RPDF");
            }

            oCampoInvoices.FromPane = 5;
            oCampoInvoices.ToPane = 5;

            #endregion

            #region Campo XML enviado 

            //*******************************************
            // Se adiciona Label "XML Enviado"
            //*******************************************

            oCampoInvoices = _oFormInvoices.Items.Add("lblXML", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oCampoInvoices.Left = oItem.Left + 10;
            oCampoInvoices.Width = oItem.Width;
            oCampoInvoices.Top = oItem.Top + 60;
            oCampoInvoices.Height = oItem.Height;

            oCampoInvoices.LinkTo = "txtXML";

            oStaticText = ((SAPbouiCOM.StaticText)(oCampoInvoices.Specific));

            oStaticText.Caption = "XML Enviado";

            oCampoInvoices.FromPane = 5;
            oCampoInvoices.ToPane = 5;

            //*******************************************
            // Se adiciona Tex Box "XML Enviado"
            //*******************************************

            oCampoInvoices = _oFormInvoices.Items.Add("txtXML", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oCampoInvoices.Left = oItem.Left + 120;
            oCampoInvoices.Width = oItem.Width + 100;
            oCampoInvoices.Top = oItem.Top + 60;
            oCampoInvoices.Height = oItem.Height;
            oCampoInvoices.Enabled = false;

            otxtRPDF = ((SAPbouiCOM.EditText)(oCampoInvoices.Specific));


            if (__TipoDoc == "FacturaDeClientes" || __TipoDoc == "NotaDebitoClientes")
            {
                otxtRPDF.DataBind.SetBound(true, "OINV", "U_BO_XML");
            }
            else if (__TipoDoc == "NotaCreditoClientes")
            {
                otxtRPDF.DataBind.SetBound(true, "ORIN", "U_BO_XML");
            }

            oCampoInvoices.FromPane = 5;
            oCampoInvoices.ToPane = 5;

            #endregion

            #region Campo Estado Documento 

            //*******************************************
            // Se adiciona Label "Estado Documento"
            //*******************************************

            oCampoInvoices = _oFormInvoices.Items.Add("lblS", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oCampoInvoices.Left = oItem.Left + 10;
            oCampoInvoices.Width = oItem.Width;
            oCampoInvoices.Top = oItem.Top + 80;
            oCampoInvoices.Height = oItem.Height;

            oCampoInvoices.LinkTo = "cboS";

            oStaticText = ((SAPbouiCOM.StaticText)(oCampoInvoices.Specific));

            oStaticText.Caption = "Estado Documento";

            oCampoInvoices.FromPane = 5;
            oCampoInvoices.ToPane = 5;

            //*******************************************
            // Se adiciona Tex Box "Estado Documento"
            //*******************************************

            oCampoInvoices = _oFormInvoices.Items.Add("cboS", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            oCampoInvoices.Left = oItem.Left + 120;
            oCampoInvoices.Width = oItem.Width;
            oCampoInvoices.Top = oItem.Top + 80;
            oCampoInvoices.Height = oItem.Height;
            oCampoInvoices.Enabled = false;

            oCampoInvoices.DisplayDesc = true;
            _oFormInvoices.DataSources.UserDataSources.Add("cboS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);

            cboS = ((SAPbouiCOM.ComboBox)(oCampoInvoices.Specific));


            if (__TipoDoc == "FacturaDeClientes" || __TipoDoc == "NotaDebitoClientes")
            {
                cboS.DataBind.SetBound(true, "OINV", "U_BO_S");
            }
            else if (__TipoDoc == "NotaCreditoClientes")
            {
                cboS.DataBind.SetBound(true, "ORIN", "U_BO_S");
            }

            cboS.ValidValues.Add("0", "A la espera");
            cboS.ValidValues.Add("1", "Aceptada");
            cboS.ValidValues.Add("2", "Rechazada");
            cboS.ValidValues.Add("3", "En Validación");

            cboS.Select("0", BoSearchKey.psk_ByValue);

            oCampoInvoices.FromPane = 5;
            oCampoInvoices.ToPane = 5;

            #endregion

            #region Campo CUFE

            //*******************************************
            // Se adiciona Label "CUFE"
            //*******************************************

            oCampoInvoices = _oFormInvoices.Items.Add("lblCUFE", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oCampoInvoices.Left = oItem.Left + 10;
            oCampoInvoices.Width = oItem.Width;
            oCampoInvoices.Top = oItem.Top + 100;
            oCampoInvoices.Height = oItem.Height;

            oCampoInvoices.LinkTo = "txtCUFE";

            oStaticText = ((SAPbouiCOM.StaticText)(oCampoInvoices.Specific));

            oStaticText.Caption = "CUFE";

            oCampoInvoices.FromPane = 5;
            oCampoInvoices.ToPane = 5;

            //*******************************************
            // Se adiciona Tex Box "CUFE"
            //*******************************************

            oCampoInvoices = _oFormInvoices.Items.Add("txtCUFE", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oCampoInvoices.Left = oItem.Left + 120;
            oCampoInvoices.Width = oItem.Width + 150;
            oCampoInvoices.Top = oItem.Top + 100;
            oCampoInvoices.Height = oItem.Height;

            otxtCUFE = ((SAPbouiCOM.EditText)(oCampoInvoices.Specific));

            if (__TipoDoc == "FacturaDeClientes" || __TipoDoc == "NotaDebitoClientes")
            {
                otxtCUFE.DataBind.SetBound(true, "OINV", "U_BO_CUFE");
            }
            else if (__TipoDoc == "NotaCreditoClientes")
            {
                otxtCUFE.DataBind.SetBound(true, "ORIN", "U_BO_CUFE");
            }

            oCampoInvoices.FromPane = 5;
            oCampoInvoices.ToPane = 5;

            #endregion

            #region Campo Enviado por 

            //*******************************************
            // Se adiciona Label "Enviado por"
            //*******************************************

            oCampoInvoices = _oFormInvoices.Items.Add("lblPP", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oCampoInvoices.Left = oItem.Left + 10;
            oCampoInvoices.Width = oItem.Width;
            oCampoInvoices.Top = oItem.Top + 120;
            oCampoInvoices.Height = oItem.Height;

            oCampoInvoices.LinkTo = "txtPP";

            oStaticText = ((SAPbouiCOM.StaticText)(oCampoInvoices.Specific));

            oStaticText.Caption = "Enviado por";

            oCampoInvoices.FromPane = 5;
            oCampoInvoices.ToPane = 5;

            //*******************************************
            // Se adiciona Tex Box "Enviado por"
            //*******************************************

            oCampoInvoices = _oFormInvoices.Items.Add("txtPP", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            oCampoInvoices.Left = oItem.Left + 120;
            oCampoInvoices.Width = oItem.Width;
            oCampoInvoices.Top = oItem.Top + 120;
            oCampoInvoices.Height = oItem.Height;
            oCampoInvoices.Enabled = false;

            oCampoInvoices.DisplayDesc = true;

            _oFormInvoices.DataSources.UserDataSources.Add("cboPP", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);

            cboPP = ((SAPbouiCOM.ComboBox)(oCampoInvoices.Specific));

            if (__TipoDoc == "FacturaDeClientes" || __TipoDoc == "NotaDebitoClientes")
            {
                cboPP.DataBind.SetBound(true, "OINV", "U_BO_PP");
            }
            else if (__TipoDoc == "NotaCreditoClientes")
            {
                cboPP.DataBind.SetBound(true, "ORIN", "U_BO_PP");
            }

            cboPP.ValidValues.Add("A", "AddIn");
            cboPP.ValidValues.Add("M", "Masivo");

            cboPP.Select("A", BoSearchKey.psk_ByValue);

            oCampoInvoices.FromPane = 5;
            oCampoInvoices.ToPane = 5;

            #endregion

            #region Campo Codigo QR

            //*******************************************
            // Se adiciona Label "Codigo QR"
            //*******************************************

            oCampoInvoices = _oFormInvoices.Items.Add("lblQR", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oCampoInvoices.Left = oItem.Left + 10;
            oCampoInvoices.Width = oItem.Width;
            oCampoInvoices.Top = oItem.Top + 140;
            oCampoInvoices.Height = oItem.Height;

            oCampoInvoices.LinkTo = "txtQR";

            oStaticText = ((SAPbouiCOM.StaticText)(oCampoInvoices.Specific));

            oStaticText.Caption = "Codigo QR";

            oCampoInvoices.FromPane = 5;
            oCampoInvoices.ToPane = 5;

            //*******************************************
            // Se adiciona Tex Box "Codigo QR"
            //*******************************************

            oCampoInvoices = _oFormInvoices.Items.Add("txtQR", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oCampoInvoices.Left = oItem.Left + 120;
            oCampoInvoices.Width = oItem.Width;
            oCampoInvoices.Top = oItem.Top + 140;
            oCampoInvoices.Height = oItem.Height;

            otxtQR = ((SAPbouiCOM.EditText)(oCampoInvoices.Specific));

            if (__TipoDoc == "FacturaDeClientes" || __TipoDoc == "NotaDebitoClientes")
            {
                otxtQR.DataBind.SetBound(true, "OINV", "U_BO_QR");
            }
            else if (__TipoDoc == "NotaCreditoClientes")
            {
                otxtQR.DataBind.SetBound(true, "ORIN", "U_BO_QR");
            }

            oCampoInvoices.FromPane = 5;
            oCampoInvoices.ToPane = 5;

            #endregion

            #region Tipo de descuento - Aplica solo FV

            //*******************************************
            // Se adiciona Label "Tipo de descuento"
            //*******************************************

            if (__TipoDoc == "FacturaDeClientes" || __TipoDoc == "NotaDebitoClientes")
            {
                oCampoInvoices = _oFormInvoices.Items.Add("lblTDesc", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oCampoInvoices.Left = oItem.Left + 270;
                oCampoInvoices.Width = oItem.Width;
                oCampoInvoices.Top = oItem.Top + 120;
                oCampoInvoices.Height = oItem.Height;

                oCampoInvoices.LinkTo = "txtTDesc";

                oStaticText = ((SAPbouiCOM.StaticText)(oCampoInvoices.Specific));

                oStaticText.Caption = "Tipo Descuento";

                oCampoInvoices.FromPane = 5;
                oCampoInvoices.ToPane = 5;

                //*******************************************
                // Se adiciona Tex Box "Tipo de descuento"
                //*******************************************

                oCampoInvoices = _oFormInvoices.Items.Add("txtTDesc", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                oCampoInvoices.Left = oItem.Left + 350;
                oCampoInvoices.Width = oItem.Width;
                oCampoInvoices.Top = oItem.Top + 120;
                oCampoInvoices.Height = oItem.Height;

                oCampoInvoices.DisplayDesc = true;

                _oFormInvoices.DataSources.UserDataSources.Add("cboTDesc", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);

                cboTDesc = ((SAPbouiCOM.ComboBox)(oCampoInvoices.Specific));

                cboTDesc.DataBind.SetBound(true, "OINV", "U_BO_DESC");

                cboPP.ValidValues.Add("00", "Descuento por impuesto asumido");
                cboPP.ValidValues.Add("01", "Pague uno lleve otro");
                cboPP.ValidValues.Add("02", "Descuentos contractuales");
                cboPP.ValidValues.Add("03", "Descuento por pronto pago");
                cboPP.ValidValues.Add("04", "Envío gratis");
                cboPP.ValidValues.Add("05", "Descuentos específicos por inventarios");
                cboPP.ValidValues.Add("06", "Descuento por monto de compras");
                cboPP.ValidValues.Add("07", "Descuento de temporada");
                cboPP.ValidValues.Add("08", "Descuento por actualización de productos / servicios");
                cboPP.ValidValues.Add("09", "Descuento general");
                cboPP.ValidValues.Add("10", "Descuento por volumen");
                cboPP.ValidValues.Add("11", "Otro descuento");

                cboPP.Select("09", BoSearchKey.psk_ByValue);

                oCampoInvoices.FromPane = 5;
                oCampoInvoices.ToPane = 5;

            }

            #endregion

            #region Campo Medio de Pago

            //*******************************************
            // Se adiciona Label "Medio de Pago"
            //*******************************************

            oCampoInvoices = _oFormInvoices.Items.Add("lblMP", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oCampoInvoices.Left = oItem.Left + 10;
            oCampoInvoices.Width = oItem.Width;
            oCampoInvoices.Top = oItem.Top + 160;
            oCampoInvoices.Height = oItem.Height;

            oCampoInvoices.LinkTo = "txtMP";

            oStaticText = ((SAPbouiCOM.StaticText)(oCampoInvoices.Specific));

            oStaticText.Caption = "Med. de Pago";

            oCampoInvoices.FromPane = 5;
            oCampoInvoices.ToPane = 5;


            //*******************************************
            // Se adiciona Tex Box "Medio de Pago"
            //*******************************************

            oCampoInvoices = _oFormInvoices.Items.Add("txtMP", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            oCampoInvoices.Left = oItem.Left + 120;
            oCampoInvoices.Width = oItem.Width;
            oCampoInvoices.Top = oItem.Top + 160;
            oCampoInvoices.Height = oItem.Height;


            oCampoInvoices.DisplayDesc = true;

            _oFormInvoices.DataSources.UserDataSources.Add("cboMP", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);

            cboMP = ((SAPbouiCOM.ComboBox)(oCampoInvoices.Specific));

            if (__TipoDoc == "FacturaDeClientes" || __TipoDoc == "NotaDebitoClientes")
            {
                cboMP.DataBind.SetBound(true, "OINV", "U_BO_MP");
            }
            else if (__TipoDoc == "NotaCreditoClientes")
            {
                cboMP.DataBind.SetBound(true, "ORIN", "U_BO_MP");
            }

            cboMP.ValidValues.Add("1", "Instrumento no definido");
            cboMP.ValidValues.Add("2", "Crédito ACH");
            cboMP.ValidValues.Add("3", "Débito ACH");
            cboMP.ValidValues.Add("4", "Reversión débito de demanda ACH");
            cboMP.ValidValues.Add("5", "Reversión crédito de demanda ACH");
            cboMP.ValidValues.Add("6", "Reversión crédito de demanda ACH");
            cboMP.ValidValues.Add("7", "Débito de demanda ACH");
            cboMP.ValidValues.Add("8", "Mantener");
            cboMP.ValidValues.Add("9", "Clearing Nacional o Regional");
            cboMP.ValidValues.Add("10", "Efectivo");
            cboMP.ValidValues.Add("11", "Reversión Crédito Ahorro");
            cboMP.ValidValues.Add("12", "Reversión Débito Ahorro");
            cboMP.ValidValues.Add("13", "Crédito Ahorro");
            cboMP.ValidValues.Add("14", "Débito Ahorro");
            cboMP.ValidValues.Add("15", "Bookentry Crédito");
            cboMP.ValidValues.Add("16", "Bookentry Débito");
            cboMP.ValidValues.Add("17", "Concentración de la demanda en efectivo / Crédito (CCD)");
            cboMP.ValidValues.Add("18", "Concentración de la demanda en efectivo / Debito (CCD)");
            cboMP.ValidValues.Add("19", "Crédito Pago negocio corporativo (CTP)");
            cboMP.ValidValues.Add("20", "Cheque");
            cboMP.ValidValues.Add("21", "Proyecto bancario");
            cboMP.ValidValues.Add("22", "Proyecto bancario certificado");
            cboMP.ValidValues.Add("23", "Cheque bancario");
            cboMP.ValidValues.Add("24", "Nota cambiaria esperando aceptación");
            cboMP.ValidValues.Add("25", "Cheque certificado");
            cboMP.ValidValues.Add("26", "Cheque local");
            cboMP.ValidValues.Add("27", "Débito Pago Negocio Corporativo (CTP)");
            cboMP.ValidValues.Add("28", "Crédito Negocio Intercambio Corporativo (CTX)");
            cboMP.ValidValues.Add("29", "Débito Negocio Intercambio Corporativo (CTX)");
            cboMP.ValidValues.Add("30", "Transferencia Crédito");
            cboMP.ValidValues.Add("31", "Transferencia Débito");
            cboMP.ValidValues.Add("32", "Concentración Efectivo / Desembolso Crédito plus");
            cboMP.ValidValues.Add("33", "Concentración Efectivo / Desembolso Débito plus");
            cboMP.ValidValues.Add("34", "Pago y depósito pre acordado");
            cboMP.ValidValues.Add("35", "Concentración efectivo");
            cboMP.ValidValues.Add("36", "Concentración efectivo ahorros / Desembolso");
            cboMP.ValidValues.Add("37", "Pago Negocio Corporativo Ahorros Crédito");
            cboMP.ValidValues.Add("38", "Pago Negocio Corporativo Ahorros Débito");
            cboMP.ValidValues.Add("39", "Crédito Negocio Intercambio Corporativo");
            cboMP.ValidValues.Add("40", "Débito Negocio Intercambio Corporativo");
            cboMP.ValidValues.Add("41", "Concentración efectivo/Desembolso Crédito plus");
            cboMP.ValidValues.Add("42", "Consignación bancaria");
            cboMP.ValidValues.Add("43", "Concentración efectivo / Desembolso Débito plus");
            cboMP.ValidValues.Add("44", "Nota cambiaria");
            cboMP.ValidValues.Add("45", "Transferencia Crédito Bancario");
            cboMP.ValidValues.Add("46", "Transferencia Débito Interbancario");
            cboMP.ValidValues.Add("47", "Transferencia Débito Bancaria");
            cboMP.ValidValues.Add("48", "Tarjeta Crédito");
            cboMP.ValidValues.Add("49", "Tarjeta Débito");
            cboMP.ValidValues.Add("50", "Pstgiro");
            cboMP.ValidValues.Add("51", "Telex estándar bancario francés");
            cboMP.ValidValues.Add("52", "Pago comercial Urgente");
            cboMP.ValidValues.Add("53", "Pago Tesorería Urgente");
            cboMP.ValidValues.Add("60", "Nota promisoria");
            cboMP.ValidValues.Add("61", "Nota promisoria firmada por el acreedor");
            cboMP.ValidValues.Add("62", "Nota promisoria firmada por el acreedor, avalada por el banco");
            cboMP.ValidValues.Add("63", "Nota promisoria firmada por el acreedor, avalada por un tercero");
            cboMP.ValidValues.Add("64", "Nota promisoria firmada por el banco");
            cboMP.ValidValues.Add("65", "Nota promisoria firmada por un banco, avalada por otro banco");
            cboMP.ValidValues.Add("66", "Nota promisoria firmada");
            cboMP.ValidValues.Add("67", "Nota promisoria firmada por un tercero avalada por un banco");
            cboMP.ValidValues.Add("70", "Retiro de nota por el acreedor");
            cboMP.ValidValues.Add("74", "Retiro de nota por el acreedor sobre un banco");
            cboMP.ValidValues.Add("75", "Retiro de nota por el acreedor, avalada por otro banco");
            cboMP.ValidValues.Add("76", "Retiro de nota por el acreedor, sobre un banco avalada por un tercero");
            cboMP.ValidValues.Add("77", "Retiro de nota por el acreedor sobre un tercero");
            cboMP.ValidValues.Add("78", "Retiro de nota por el acreedor sobre un tercero avalada por un banco");
            cboMP.ValidValues.Add("91", "Nota bancaria transferible");
            cboMP.ValidValues.Add("92", "Cheque local transferible");
            cboMP.ValidValues.Add("93", "Giro referenciado");
            cboMP.ValidValues.Add("94", "Giro Urgente");
            cboMP.ValidValues.Add("95", "Giro formato abierto");
            cboMP.ValidValues.Add("96", "Método de pago solicitado no usado");
            cboMP.ValidValues.Add("97", "Clearing entre partners");
            cboMP.ValidValues.Add("ZZZ", "Acuerdo mutuo");

            cboMP.Select("10", BoSearchKey.psk_ByValue);

            oCampoInvoices.FromPane = 5;
            oCampoInvoices.ToPane = 5;

            #endregion

            #region Campo Aplica a Factura de venta - Solo para NC y ND

            if (__TipoDoc == "NotaCreditoClientes" || __TipoDoc == "NotaDebitoClientes")
            {
                //*******************************************
                // Se adiciona Label "Aplica a Factura de venta - Solo para Notas Credito"
                //*******************************************

                oCampoInvoices = _oFormInvoices.Items.Add("blbAFV", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oCampoInvoices.Left = oItem.Left + 10;
                oCampoInvoices.Width = oItem.Width;
                oCampoInvoices.Top = oItem.Top + 180;
                oCampoInvoices.Height = oItem.Height;

                oCampoInvoices.LinkTo = "txtAFV";

                oStaticText = ((SAPbouiCOM.StaticText)(oCampoInvoices.Specific));

                oStaticText.Caption = "Aplicar a FV No.";

                oCampoInvoices.FromPane = 5;
                oCampoInvoices.ToPane = 5;

                //*******************************************
                // Se adiciona Tex Box "Comentarios Fac.Elec"
                //*******************************************

                oCampoInvoices = _oFormInvoices.Items.Add("txtAFV", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oCampoInvoices.Left = oItem.Left + 120;
                oCampoInvoices.Width = oItem.Width;
                oCampoInvoices.Top = oItem.Top + 180;
                oCampoInvoices.Height = oItem.Height;

                otxtAFV = ((SAPbouiCOM.EditText)(oCampoInvoices.Specific));

                if (__TipoDoc == "NotaCreditoClientes")
                {
                    otxtAFV.DataBind.SetBound(true, "ORIN", "U_BO_AFV");
                }
                else if (__TipoDoc == "NotaDebitoClientes")
                {
                    otxtAFV.DataBind.SetBound(true, "OINV", "U_BO_AFV");
                }
            }

            if (__TipoDoc == "FacturaDeClientes")
            {
                otxtCEB.DataBind.SetBound(true, "OINV", "U_BO_EBC");
            }
            else if (__TipoDoc == "NotaCreditoClientes")
            {
                otxtCEB.DataBind.SetBound(true, "ORIN", "U_BO_EBC");
            }

            oCampoInvoices.FromPane = 5;
            oCampoInvoices.ToPane = 5;

            #endregion

            #region Tipo de Nota - Aplica solo NC

            if (__TipoDoc == "NotaCreditoClientes")
            {
                //*******************************************
                // Se adiciona Label "Tipo de Nota - Solo para Notas Credito"
                //*******************************************

                oCampoInvoices = _oFormInvoices.Items.Add("blbTipN", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oCampoInvoices.Left = oItem.Left + 10;
                oCampoInvoices.Width = oItem.Width;
                oCampoInvoices.Top = oItem.Top + 200;
                oCampoInvoices.Height = oItem.Height;

                oCampoInvoices.LinkTo = "txtTipN";

                oStaticText = ((SAPbouiCOM.StaticText)(oCampoInvoices.Specific));

                oStaticText.Caption = "Tipo Nota.";

                oCampoInvoices.FromPane = 5;
                oCampoInvoices.ToPane = 5;

                //*******************************************
                // Se adiciona Tex Box "Tipo de Nota"
                //*******************************************

                oCampoInvoices = _oFormInvoices.Items.Add("txtTipN", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                oCampoInvoices.Left = oItem.Left + 120;
                oCampoInvoices.Width = oItem.Width;
                oCampoInvoices.Top = oItem.Top + 200;
                oCampoInvoices.Height = oItem.Height;

                oCampoInvoices.DisplayDesc = true;

                otxtTN = ((SAPbouiCOM.ComboBox)(oCampoInvoices.Specific));

                otxtTN.DataBind.SetBound(true, "ORIN", "U_BO_TN");

                otxtTN.ValidValues.Add("1", "Devolución de parte de los bienes");
                otxtTN.ValidValues.Add("2", "Anulación de factura electrónica");
                otxtTN.ValidValues.Add("3", "Rebaja");
                otxtTN.ValidValues.Add("4", "Descuento");


                otxtTN.Select("1", BoSearchKey.psk_ByValue);

                oCampoInvoices.FromPane = 5;
                oCampoInvoices.ToPane = 5;

            }



            #endregion

            #region Tipo de Nota - Aplica solo ND

            if (__TipoDoc == "NotaDebitoClientes")
            {
                //*******************************************
                // Se adiciona Label "Tipo de Nota - Solo para Notas Credito"
                //*******************************************

                oCampoInvoices = _oFormInvoices.Items.Add("blbTipND", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oCampoInvoices.Left = oItem.Left + 10;
                oCampoInvoices.Width = oItem.Width;
                oCampoInvoices.Top = oItem.Top + 200;
                oCampoInvoices.Height = oItem.Height;

                oCampoInvoices.LinkTo = "txtTipND";

                oStaticText = ((SAPbouiCOM.StaticText)(oCampoInvoices.Specific));

                oStaticText.Caption = "Tipo Nota Debito.";

                oCampoInvoices.FromPane = 5;
                oCampoInvoices.ToPane = 5;

                //*******************************************
                // Se adiciona Tex Box "Tipo de Nota"
                //*******************************************

                oCampoInvoices = _oFormInvoices.Items.Add("txtTipND", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                oCampoInvoices.Left = oItem.Left + 120;
                oCampoInvoices.Width = oItem.Width;
                oCampoInvoices.Top = oItem.Top + 200;
                oCampoInvoices.Height = oItem.Height;

                oCampoInvoices.DisplayDesc = true;

                otxtTND = ((SAPbouiCOM.ComboBox)(oCampoInvoices.Specific));

                otxtTND.DataBind.SetBound(true, "OINV", "U_BO_TipND");

                otxtTND.ValidValues.Add("1", "Intereses");
                otxtTND.ValidValues.Add("2", "Gastos por cobrar");
                otxtTND.ValidValues.Add("3", "Cambio de valor");
                otxtTND.ValidValues.Add("4", "Otro");


                otxtTND.Select("3", BoSearchKey.psk_ByValue);

                oCampoInvoices.FromPane = 5;
                oCampoInvoices.ToPane = 5;

            }



            #endregion

            #region Vacia Campós en Notas Credito

            if (_oFormInvoices.Mode == BoFormMode.fm_ADD_MODE && __TipoDoc == "NotaCreditoClientes")
            {
                #region Variables

                string ValorGeneral = "_";

                SAPbouiCOM.EditText _txtCRWS = (SAPbouiCOM.EditText)(_oFormInvoices.Items.Item("txtCRWS").Specific);
                SAPbouiCOM.EditText _txtMRWS = (SAPbouiCOM.EditText)(_oFormInvoices.Items.Item("txtMRWS").Specific);
                SAPbouiCOM.EditText _txtRPDF = (SAPbouiCOM.EditText)(_oFormInvoices.Items.Item("txtRPDF").Specific);
                SAPbouiCOM.EditText _txtXML = (SAPbouiCOM.EditText)(_oFormInvoices.Items.Item("txtXML").Specific);
                SAPbouiCOM.EditText _txtCUFE = (SAPbouiCOM.EditText)(_oFormInvoices.Items.Item("txtCUFE").Specific);
                SAPbouiCOM.EditText _txtQR = (SAPbouiCOM.EditText)(_oFormInvoices.Items.Item("txtQR").Specific);

                #endregion

                _txtCRWS.Value = ValorGeneral;
                _txtMRWS.Value = ValorGeneral;
                _txtRPDF.Value = ValorGeneral;
                _txtXML.Value = ValorGeneral;
                _txtCUFE.Value = ValorGeneral;
                _txtQR.Value = ValorGeneral;


            }

            #endregion

        }

        private void ItemsBusinessParnerd(SAPbouiCOM.Form oFormBusinessParnerd, int _QuantityEmails)
        {
            SAPbouiCOM.Form _oFormBusinessParnerd;
            SAPbouiCOM.Item oItem;
            SAPbouiCOM.Item oCampoInvoices = null;
            SAPbouiCOM.StaticText oStaticText = null;

            SAPbouiCOM.EditText otxtEmail1 = null;
            SAPbouiCOM.ComboBox cboTR = null;

            _oFormBusinessParnerd = oFormBusinessParnerd;

            oItem = _oFormBusinessParnerd.Items.Item("44");

            if (_QuantityEmails == 1)
            {
                #region Adiciona 1 label y texbox para los correos

                //*******************************************
                // Se adiciona Label "Correo Electronico 1"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("blbEmail1", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oCampoInvoices.Left = oItem.Left + 10;
                oCampoInvoices.Width = oItem.Width;
                oCampoInvoices.Top = oItem.Top + 10;
                oCampoInvoices.Height = oItem.Height;

                oCampoInvoices.LinkTo = "txtEmail1";

                oStaticText = ((SAPbouiCOM.StaticText)(oCampoInvoices.Specific));

                oStaticText.Caption = "Correo Electronico 1 ";

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                //*******************************************
                // Se adiciona Tex Box "Correo Electronico 1"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("txtEmail1", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oCampoInvoices.Left = oItem.Left + 120;
                oCampoInvoices.Width = oItem.Width + 50;
                oCampoInvoices.Top = oItem.Top + 10;
                oCampoInvoices.Height = oItem.Height;

                otxtEmail1 = ((SAPbouiCOM.EditText)(oCampoInvoices.Specific));

                otxtEmail1.DataBind.SetBound(true, "OCRD", "U_BO_E_mail_1");

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                //*******************************************
                // Se adiciona Label "Respon. Fiscal"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("blbRF", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oCampoInvoices.Left = oItem.Left + 10;
                oCampoInvoices.Width = oItem.Width;
                oCampoInvoices.Top = oItem.Top + 30;
                oCampoInvoices.Height = oItem.Height;

                oCampoInvoices.LinkTo = "txtRF";

                oStaticText = ((SAPbouiCOM.StaticText)(oCampoInvoices.Specific));

                oStaticText.Caption = "Respon. Fiscal";

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                //*******************************************
                // Se adiciona Tex Box "Respon. Fiscal"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("txtRF", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oCampoInvoices.Left = oItem.Left + 120;
                oCampoInvoices.Width = oItem.Width + 50;
                oCampoInvoices.Top = oItem.Top + 30;
                oCampoInvoices.Height = oItem.Height;

                otxtEmail1 = ((SAPbouiCOM.EditText)(oCampoInvoices.Specific));

                otxtEmail1.DataBind.SetBound(true, "OCRD", "U_BO_RF");

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                //*******************************************
                // Se adiciona Label "Tipos de Regimen"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("lblTR", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oCampoInvoices.Left = oItem.Left + 10;
                oCampoInvoices.Width = oItem.Width;
                oCampoInvoices.Top = oItem.Top + 50;
                oCampoInvoices.Height = oItem.Height;

                oCampoInvoices.LinkTo = "cboTR";

                oStaticText = ((SAPbouiCOM.StaticText)(oCampoInvoices.Specific));

                oStaticText.Caption = "Tipos de Regimen";

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                //*******************************************
                // Se adiciona Tex Box "Tipos de Regimen"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("cboTR", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                oCampoInvoices.Left = oItem.Left + 120;
                oCampoInvoices.Width = oItem.Width;
                oCampoInvoices.Top = oItem.Top + 50;
                oCampoInvoices.Height = oItem.Height;
                oCampoInvoices.Enabled = true;

                oCampoInvoices.DisplayDesc = true;
                _oFormBusinessParnerd.DataSources.UserDataSources.Add("cboTR", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);

                cboTR = ((SAPbouiCOM.ComboBox)(oCampoInvoices.Specific));

                cboTR.DataBind.SetBound(true, "OCRD", "U_BO_TR");

                cboTR.ValidValues.Add("04", "Régimen Simple");
                cboTR.ValidValues.Add("05", "Régimen Ordinario");
                cboTR.ValidValues.Add("48", "Impuesto sobre las ventas - IVA");
                cboTR.ValidValues.Add("49", "No responsable de IVA");

                //cboTR.Select("0", BoSearchKey.psk_ByValue);

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                #endregion
            }
            else if (_QuantityEmails == 2)
            {
                #region Adiciona 2 label y texbox para los correos

                //*******************************************
                // Se adiciona Label "Correo Electronico 1"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("blbEmail1", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oCampoInvoices.Left = oItem.Left + 10;
                oCampoInvoices.Width = oItem.Width;
                oCampoInvoices.Top = oItem.Top + 10;
                oCampoInvoices.Height = oItem.Height;

                oCampoInvoices.LinkTo = "txtEmail1";

                oStaticText = ((SAPbouiCOM.StaticText)(oCampoInvoices.Specific));

                oStaticText.Caption = "Correo Electronico 1 ";

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                //*******************************************
                // Se adiciona Tex Box "Correo Electronico 1"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("txtEmail1", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oCampoInvoices.Left = oItem.Left + 120;
                oCampoInvoices.Width = oItem.Width + 50;
                oCampoInvoices.Top = oItem.Top + 10;
                oCampoInvoices.Height = oItem.Height;

                otxtEmail1 = ((SAPbouiCOM.EditText)(oCampoInvoices.Specific));

                otxtEmail1.DataBind.SetBound(true, "OCRD", "U_BO_E_mail_1");

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                //*******************************************
                // Se adiciona Label "Correo Electronico 2"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("blbEmail2", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oCampoInvoices.Left = oItem.Left + 10;
                oCampoInvoices.Width = oItem.Width;
                oCampoInvoices.Top = oItem.Top + 30;
                oCampoInvoices.Height = oItem.Height;

                oCampoInvoices.LinkTo = "txtEmail2";

                oStaticText = ((SAPbouiCOM.StaticText)(oCampoInvoices.Specific));

                oStaticText.Caption = "Correo Electronico 2 ";

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                //*******************************************
                // Se adiciona Tex Box "Correo Electronico 2"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("txtEmail2", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oCampoInvoices.Left = oItem.Left + 120;
                oCampoInvoices.Width = oItem.Width + 50;
                oCampoInvoices.Top = oItem.Top + 30;
                oCampoInvoices.Height = oItem.Height;

                otxtEmail1 = ((SAPbouiCOM.EditText)(oCampoInvoices.Specific));

                otxtEmail1.DataBind.SetBound(true, "OCRD", "U_BO_E_mail_2");

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                //*******************************************
                // Se adiciona Label "Respon. Fiscal"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("blbRF", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oCampoInvoices.Left = oItem.Left + 10;
                oCampoInvoices.Width = oItem.Width;
                oCampoInvoices.Top = oItem.Top + 50;
                oCampoInvoices.Height = oItem.Height;

                oCampoInvoices.LinkTo = "txtRF";

                oStaticText = ((SAPbouiCOM.StaticText)(oCampoInvoices.Specific));

                oStaticText.Caption = "Respon. Fiscal";

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                //*******************************************
                // Se adiciona Tex Box "Respon. Fiscal"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("txtRF", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oCampoInvoices.Left = oItem.Left + 120;
                oCampoInvoices.Width = oItem.Width + 50;
                oCampoInvoices.Top = oItem.Top + 50;
                oCampoInvoices.Height = oItem.Height;

                otxtEmail1 = ((SAPbouiCOM.EditText)(oCampoInvoices.Specific));

                otxtEmail1.DataBind.SetBound(true, "OCRD", "U_BO_RF");

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                //*******************************************
                // Se adiciona Label "Tipos de Regimen"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("lblTR", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oCampoInvoices.Left = oItem.Left + 10;
                oCampoInvoices.Width = oItem.Width;
                oCampoInvoices.Top = oItem.Top + 70;
                oCampoInvoices.Height = oItem.Height;

                oCampoInvoices.LinkTo = "cboTR";

                oStaticText = ((SAPbouiCOM.StaticText)(oCampoInvoices.Specific));

                oStaticText.Caption = "Tipos de Regimen";

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                //*******************************************
                // Se adiciona Tex Box "Tipos de Regimen"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("cboTR", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                oCampoInvoices.Left = oItem.Left + 120;
                oCampoInvoices.Width = oItem.Width;
                oCampoInvoices.Top = oItem.Top + 70;
                oCampoInvoices.Height = oItem.Height;

                oCampoInvoices.DisplayDesc = true;
                _oFormBusinessParnerd.DataSources.UserDataSources.Add("cboTR", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);

                cboTR = ((SAPbouiCOM.ComboBox)(oCampoInvoices.Specific));

                cboTR.DataBind.SetBound(true, "OCRD", "U_BO_TR");

                cboTR.ValidValues.Add("04", "Régimen Simple");
                cboTR.ValidValues.Add("05", "Régimen Ordinario");
                cboTR.ValidValues.Add("48", "Impuesto sobre las ventas - IVA");
                cboTR.ValidValues.Add("49", "No responsable de IVA");



                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                #endregion
            }
            else if (_QuantityEmails == 3)
            {
                #region Adiciona 3 label y texbox para los correos

                //*******************************************
                // Se adiciona Label "Correo Electronico 1"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("blbEmail1", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oCampoInvoices.Left = oItem.Left + 10;
                oCampoInvoices.Width = oItem.Width;
                oCampoInvoices.Top = oItem.Top + 10;
                oCampoInvoices.Height = oItem.Height;

                oCampoInvoices.LinkTo = "txtEmail1";

                oStaticText = ((SAPbouiCOM.StaticText)(oCampoInvoices.Specific));

                oStaticText.Caption = "Correo Electronico 1 ";

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                //*******************************************
                // Se adiciona Tex Box "Correo Electronico 1"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("txtEmail1", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oCampoInvoices.Left = oItem.Left + 120;
                oCampoInvoices.Width = oItem.Width + 50;
                oCampoInvoices.Top = oItem.Top + 10;
                oCampoInvoices.Height = oItem.Height;

                otxtEmail1 = ((SAPbouiCOM.EditText)(oCampoInvoices.Specific));

                otxtEmail1.DataBind.SetBound(true, "OCRD", "U_BO_E_mail_1");

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                //*******************************************
                // Se adiciona Label "Correo Electronico 2"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("blbEmail2", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oCampoInvoices.Left = oItem.Left + 10;
                oCampoInvoices.Width = oItem.Width;
                oCampoInvoices.Top = oItem.Top + 30;
                oCampoInvoices.Height = oItem.Height;

                oCampoInvoices.LinkTo = "txtEmail2";

                oStaticText = ((SAPbouiCOM.StaticText)(oCampoInvoices.Specific));

                oStaticText.Caption = "Correo Electronico 2 ";

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                //*******************************************
                // Se adiciona Tex Box "Correo Electronico 2"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("txtEmail2", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oCampoInvoices.Left = oItem.Left + 120;
                oCampoInvoices.Width = oItem.Width + 50;
                oCampoInvoices.Top = oItem.Top + 30;
                oCampoInvoices.Height = oItem.Height;

                otxtEmail1 = ((SAPbouiCOM.EditText)(oCampoInvoices.Specific));

                otxtEmail1.DataBind.SetBound(true, "OCRD", "U_BO_E_mail_2");

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                //*******************************************
                // Se adiciona Label "Correo Electronico 3"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("blbEmail3", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oCampoInvoices.Left = oItem.Left + 10;
                oCampoInvoices.Width = oItem.Width;
                oCampoInvoices.Top = oItem.Top + 50;
                oCampoInvoices.Height = oItem.Height;

                oCampoInvoices.LinkTo = "txtEmail3";

                oStaticText = ((SAPbouiCOM.StaticText)(oCampoInvoices.Specific));

                oStaticText.Caption = "Correo Electronico 3";

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                //*******************************************
                // Se adiciona Tex Box "Correo Electronico 3"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("txtEmail3", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oCampoInvoices.Left = oItem.Left + 120;
                oCampoInvoices.Width = oItem.Width + 50;
                oCampoInvoices.Top = oItem.Top + 50;
                oCampoInvoices.Height = oItem.Height;

                otxtEmail1 = ((SAPbouiCOM.EditText)(oCampoInvoices.Specific));

                otxtEmail1.DataBind.SetBound(true, "OCRD", "U_BO_E_mail_3");

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                //*******************************************
                // Se adiciona Label "Respon. Fiscal"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("blbRF", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oCampoInvoices.Left = oItem.Left + 10;
                oCampoInvoices.Width = oItem.Width;
                oCampoInvoices.Top = oItem.Top + 70;
                oCampoInvoices.Height = oItem.Height;

                oCampoInvoices.LinkTo = "txtRF";

                oStaticText = ((SAPbouiCOM.StaticText)(oCampoInvoices.Specific));

                oStaticText.Caption = "Respon. Fiscal";

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                //*******************************************
                // Se adiciona Tex Box "Respon. Fiscal"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("txtRF", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oCampoInvoices.Left = oItem.Left + 120;
                oCampoInvoices.Width = oItem.Width + 50;
                oCampoInvoices.Top = oItem.Top + 70;
                oCampoInvoices.Height = oItem.Height;

                otxtEmail1 = ((SAPbouiCOM.EditText)(oCampoInvoices.Specific));

                otxtEmail1.DataBind.SetBound(true, "OCRD", "U_BO_RF");

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                //*******************************************
                // Se adiciona Label "Tipos de Regimen"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("lblTR", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oCampoInvoices.Left = oItem.Left + 10;
                oCampoInvoices.Width = oItem.Width;
                oCampoInvoices.Top = oItem.Top + 90;
                oCampoInvoices.Height = oItem.Height;

                oCampoInvoices.LinkTo = "cboTR";

                oStaticText = ((SAPbouiCOM.StaticText)(oCampoInvoices.Specific));

                oStaticText.Caption = "Tipo de Regimen";

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                //*******************************************
                // Se adiciona Tex Box "Tipos de Regimen"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("cboTR", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                oCampoInvoices.Left = oItem.Left + 120;
                oCampoInvoices.Width = oItem.Width;
                oCampoInvoices.Top = oItem.Top + 90;
                oCampoInvoices.Height = oItem.Height;
                oCampoInvoices.Enabled = false;

                oCampoInvoices.DisplayDesc = true;
                _oFormBusinessParnerd.DataSources.UserDataSources.Add("cboTR", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);

                cboTR = ((SAPbouiCOM.ComboBox)(oCampoInvoices.Specific));

                cboTR.DataBind.SetBound(true, "OCRD", "U_BO_TR");

                cboTR.ValidValues.Add("04", "Régimen Simple");
                cboTR.ValidValues.Add("05", "Régimen Ordinario");
                cboTR.ValidValues.Add("48", "Impuesto sobre las ventas - IVA");
                cboTR.ValidValues.Add("49", "No responsable de IVA");

                //cboTR.Select("04", BoSearchKey.psk_ByValue);

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                #endregion
            }
            else if (_QuantityEmails == 4)
            {
                #region Adiciona 4 label y texbox para los correos

                //*******************************************
                // Se adiciona Label "Correo Electronico 1"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("blbEmail1", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oCampoInvoices.Left = oItem.Left + 10;
                oCampoInvoices.Width = oItem.Width;
                oCampoInvoices.Top = oItem.Top + 10;
                oCampoInvoices.Height = oItem.Height;

                oCampoInvoices.LinkTo = "txtEmail1";

                oStaticText = ((SAPbouiCOM.StaticText)(oCampoInvoices.Specific));

                oStaticText.Caption = "Correo Electronico 1 ";

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                //*******************************************
                // Se adiciona Tex Box "Correo Electronico 1"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("txtEmail1", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oCampoInvoices.Left = oItem.Left + 120;
                oCampoInvoices.Width = oItem.Width + 50;
                oCampoInvoices.Top = oItem.Top + 10;
                oCampoInvoices.Height = oItem.Height;

                otxtEmail1 = ((SAPbouiCOM.EditText)(oCampoInvoices.Specific));

                otxtEmail1.DataBind.SetBound(true, "OCRD", "U_BO_E_mail_1");

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                //*******************************************
                // Se adiciona Label "Correo Electronico 2"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("blbEmail2", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oCampoInvoices.Left = oItem.Left + 10;
                oCampoInvoices.Width = oItem.Width;
                oCampoInvoices.Top = oItem.Top + 30;
                oCampoInvoices.Height = oItem.Height;

                oCampoInvoices.LinkTo = "txtEmail2";

                oStaticText = ((SAPbouiCOM.StaticText)(oCampoInvoices.Specific));

                oStaticText.Caption = "Correo Electronico 2 ";

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                //*******************************************
                // Se adiciona Tex Box "Correo Electronico 2"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("txtEmail2", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oCampoInvoices.Left = oItem.Left + 120;
                oCampoInvoices.Width = oItem.Width + 50;
                oCampoInvoices.Top = oItem.Top + 30;
                oCampoInvoices.Height = oItem.Height;

                otxtEmail1 = ((SAPbouiCOM.EditText)(oCampoInvoices.Specific));

                otxtEmail1.DataBind.SetBound(true, "OCRD", "U_BO_E_mail_2");

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                //*******************************************
                // Se adiciona Label "Correo Electronico 3"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("blbEmail3", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oCampoInvoices.Left = oItem.Left + 10;
                oCampoInvoices.Width = oItem.Width;
                oCampoInvoices.Top = oItem.Top + 50;
                oCampoInvoices.Height = oItem.Height;

                oCampoInvoices.LinkTo = "txtEmail3";

                oStaticText = ((SAPbouiCOM.StaticText)(oCampoInvoices.Specific));

                oStaticText.Caption = "Correo Electronico 3";

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                //*******************************************
                // Se adiciona Tex Box "Correo Electronico 3"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("txtEmail3", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oCampoInvoices.Left = oItem.Left + 120;
                oCampoInvoices.Width = oItem.Width + 50;
                oCampoInvoices.Top = oItem.Top + 50;
                oCampoInvoices.Height = oItem.Height;

                otxtEmail1 = ((SAPbouiCOM.EditText)(oCampoInvoices.Specific));

                otxtEmail1.DataBind.SetBound(true, "OCRD", "U_BO_E_mail_3");

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                //*******************************************
                // Se adiciona Label "Correo Electronico 4"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("blbEmail4", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oCampoInvoices.Left = oItem.Left + 10;
                oCampoInvoices.Width = oItem.Width;
                oCampoInvoices.Top = oItem.Top + 70;
                oCampoInvoices.Height = oItem.Height;

                oCampoInvoices.LinkTo = "txtEmail4";

                oStaticText = ((SAPbouiCOM.StaticText)(oCampoInvoices.Specific));

                oStaticText.Caption = "Correo Electronico 4";

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                //*******************************************
                // Se adiciona Tex Box "Correo Electronico 4"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("txtEmail4", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oCampoInvoices.Left = oItem.Left + 120;
                oCampoInvoices.Width = oItem.Width + 50;
                oCampoInvoices.Top = oItem.Top + 70;
                oCampoInvoices.Height = oItem.Height;

                otxtEmail1 = ((SAPbouiCOM.EditText)(oCampoInvoices.Specific));

                otxtEmail1.DataBind.SetBound(true, "OCRD", "U_BO_E_mail_4");

                //*******************************************
                // Se adiciona Label "Respon. Fiscal"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("blbRF", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oCampoInvoices.Left = oItem.Left + 10;
                oCampoInvoices.Width = oItem.Width;
                oCampoInvoices.Top = oItem.Top + 90;
                oCampoInvoices.Height = oItem.Height;

                oCampoInvoices.LinkTo = "txtRF";

                oStaticText = ((SAPbouiCOM.StaticText)(oCampoInvoices.Specific));

                oStaticText.Caption = "Respon. Fiscal";

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                //*******************************************
                // Se adiciona Tex Box "Respon. Fiscal"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("txtRF", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oCampoInvoices.Left = oItem.Left + 120;
                oCampoInvoices.Width = oItem.Width + 50;
                oCampoInvoices.Top = oItem.Top + 90;
                oCampoInvoices.Height = oItem.Height;

                otxtEmail1 = ((SAPbouiCOM.EditText)(oCampoInvoices.Specific));

                otxtEmail1.DataBind.SetBound(true, "OCRD", "U_BO_RF");

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                //*******************************************
                // Se adiciona Label "Tipos de Regimen"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("lblTR", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oCampoInvoices.Left = oItem.Left + 10;
                oCampoInvoices.Width = oItem.Width;
                oCampoInvoices.Top = oItem.Top + 110;
                oCampoInvoices.Height = oItem.Height;

                oCampoInvoices.LinkTo = "cboTR";

                oStaticText = ((SAPbouiCOM.StaticText)(oCampoInvoices.Specific));

                oStaticText.Caption = "Tipos de Regimen";

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                //*******************************************
                // Se adiciona Tex Box "Tipos de Regimen"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("cboTR", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                oCampoInvoices.Left = oItem.Left + 120;
                oCampoInvoices.Width = oItem.Width;
                oCampoInvoices.Top = oItem.Top + 110;
                oCampoInvoices.Height = oItem.Height;
                oCampoInvoices.Enabled = false;

                oCampoInvoices.DisplayDesc = true;
                _oFormBusinessParnerd.DataSources.UserDataSources.Add("cboTR", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);

                cboTR = ((SAPbouiCOM.ComboBox)(oCampoInvoices.Specific));

                cboTR.DataBind.SetBound(true, "OCRD", "U_BO_TR");

                cboTR.ValidValues.Add("04", "Régimen Simple");
                cboTR.ValidValues.Add("05", "Régimen Ordinario");
                cboTR.ValidValues.Add("48", "Impuesto sobre las ventas - IVA");
                cboTR.ValidValues.Add("49", "No responsable de IVA");

                cboTR.Select("0", BoSearchKey.psk_ByValue);

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                #endregion
            }
            else if (_QuantityEmails == 5)
            {
                #region Adiciona 5 label y texbox para los correos

                //*******************************************
                // Se adiciona Label "Correo Electronico 1"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("blbEmail1", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oCampoInvoices.Left = oItem.Left + 10;
                oCampoInvoices.Width = oItem.Width;
                oCampoInvoices.Top = oItem.Top + 10;
                oCampoInvoices.Height = oItem.Height;

                oCampoInvoices.LinkTo = "txtEmail1";

                oStaticText = ((SAPbouiCOM.StaticText)(oCampoInvoices.Specific));

                oStaticText.Caption = "Correo Electronico 1 ";

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                //*******************************************
                // Se adiciona Tex Box "Correo Electronico 1"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("txtEmail1", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oCampoInvoices.Left = oItem.Left + 120;
                oCampoInvoices.Width = oItem.Width + 50;
                oCampoInvoices.Top = oItem.Top + 10;
                oCampoInvoices.Height = oItem.Height;

                otxtEmail1 = ((SAPbouiCOM.EditText)(oCampoInvoices.Specific));

                otxtEmail1.DataBind.SetBound(true, "OCRD", "U_BO_E_mail_1");

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                //*******************************************
                // Se adiciona Label "Correo Electronico 2"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("blbEmail2", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oCampoInvoices.Left = oItem.Left + 10;
                oCampoInvoices.Width = oItem.Width;
                oCampoInvoices.Top = oItem.Top + 30;
                oCampoInvoices.Height = oItem.Height;

                oCampoInvoices.LinkTo = "txtEmail2";

                oStaticText = ((SAPbouiCOM.StaticText)(oCampoInvoices.Specific));

                oStaticText.Caption = "Correo Electronico 2 ";

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                //*******************************************
                // Se adiciona Tex Box "Correo Electronico 2"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("txtEmail2", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oCampoInvoices.Left = oItem.Left + 120;
                oCampoInvoices.Width = oItem.Width + 50;
                oCampoInvoices.Top = oItem.Top + 30;
                oCampoInvoices.Height = oItem.Height;

                otxtEmail1 = ((SAPbouiCOM.EditText)(oCampoInvoices.Specific));

                otxtEmail1.DataBind.SetBound(true, "OCRD", "U_BO_E_mail_2");

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                //*******************************************
                // Se adiciona Label "Correo Electronico 3"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("blbEmail3", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oCampoInvoices.Left = oItem.Left + 10;
                oCampoInvoices.Width = oItem.Width;
                oCampoInvoices.Top = oItem.Top + 50;
                oCampoInvoices.Height = oItem.Height;

                oCampoInvoices.LinkTo = "txtEmail3";

                oStaticText = ((SAPbouiCOM.StaticText)(oCampoInvoices.Specific));

                oStaticText.Caption = "Correo Electronico 3";

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                //*******************************************
                // Se adiciona Tex Box "Correo Electronico 3"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("txtEmail3", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oCampoInvoices.Left = oItem.Left + 120;
                oCampoInvoices.Width = oItem.Width + 50;
                oCampoInvoices.Top = oItem.Top + 50;
                oCampoInvoices.Height = oItem.Height;

                otxtEmail1 = ((SAPbouiCOM.EditText)(oCampoInvoices.Specific));

                otxtEmail1.DataBind.SetBound(true, "OCRD", "U_BO_E_mail_3");

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                //*******************************************
                // Se adiciona Label "Correo Electronico 4"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("blbEmail4", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oCampoInvoices.Left = oItem.Left + 10;
                oCampoInvoices.Width = oItem.Width;
                oCampoInvoices.Top = oItem.Top + 70;
                oCampoInvoices.Height = oItem.Height;

                oCampoInvoices.LinkTo = "txtEmail4";

                oStaticText = ((SAPbouiCOM.StaticText)(oCampoInvoices.Specific));

                oStaticText.Caption = "Correo Electronico 4";

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                //*******************************************
                // Se adiciona Tex Box "Correo Electronico 4"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("txtEmail4", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oCampoInvoices.Left = oItem.Left + 120;
                oCampoInvoices.Width = oItem.Width + 50;
                oCampoInvoices.Top = oItem.Top + 70;
                oCampoInvoices.Height = oItem.Height;

                otxtEmail1 = ((SAPbouiCOM.EditText)(oCampoInvoices.Specific));

                otxtEmail1.DataBind.SetBound(true, "OCRD", "U_BO_E_mail_4");

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                //*******************************************
                // Se adiciona Label "Correo Electronico 5"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("blbEmail5", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oCampoInvoices.Left = oItem.Left + 10;
                oCampoInvoices.Width = oItem.Width;
                oCampoInvoices.Top = oItem.Top + 90;
                oCampoInvoices.Height = oItem.Height;

                oCampoInvoices.LinkTo = "txtEmail5";

                oStaticText = ((SAPbouiCOM.StaticText)(oCampoInvoices.Specific));

                oStaticText.Caption = "Correo Electronico 5";

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                //*******************************************
                // Se adiciona Tex Box "Correo Electronico 5"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("txtEmail5", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oCampoInvoices.Left = oItem.Left + 120;
                oCampoInvoices.Width = oItem.Width + 50;
                oCampoInvoices.Top = oItem.Top + 90;
                oCampoInvoices.Height = oItem.Height;

                otxtEmail1 = ((SAPbouiCOM.EditText)(oCampoInvoices.Specific));

                otxtEmail1.DataBind.SetBound(true, "OCRD", "U_BO_E_mail_5");

                //*******************************************
                // Se adiciona Label "Respon. Fiscal"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("blbRF", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oCampoInvoices.Left = oItem.Left + 10;
                oCampoInvoices.Width = oItem.Width;
                oCampoInvoices.Top = oItem.Top + 110;
                oCampoInvoices.Height = oItem.Height;

                oCampoInvoices.LinkTo = "txtRF";

                oStaticText = ((SAPbouiCOM.StaticText)(oCampoInvoices.Specific));

                oStaticText.Caption = "Respon. Fiscal";

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;

                //*******************************************
                // Se adiciona Tex Box "Respon. Fiscal"
                //*******************************************

                oCampoInvoices = _oFormBusinessParnerd.Items.Add("txtRF", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oCampoInvoices.Left = oItem.Left + 120;
                oCampoInvoices.Width = oItem.Width + 50;
                oCampoInvoices.Top = oItem.Top + 110;
                oCampoInvoices.Height = oItem.Height;

                otxtEmail1 = ((SAPbouiCOM.EditText)(oCampoInvoices.Specific));

                otxtEmail1.DataBind.SetBound(true, "OCRD", "U_BO_RF");

                oCampoInvoices.FromPane = 28;
                oCampoInvoices.ToPane = 28;






                #endregion
            }
        }

        public void AddItemsToDocumets(SAPbouiCOM.Form oFormInvoices, string _TipoDoc)
        {
            SAPbouiCOM.Form _oFormInvoices;
            SAPbouiCOM.Item _oNewItem;
            SAPbouiCOM.Item _oItem;
            SAPbouiCOM.Folder _oFolderItem;

            _oFormInvoices = oFormInvoices;
            _oNewItem = _oFormInvoices.Items.Add("FolderBO1", SAPbouiCOM.BoFormItemTypes.it_FOLDER);
            _oItem = _oFormInvoices.Items.Item("1320002137");

            _oNewItem.Top = _oItem.Top;
            _oNewItem.Height = _oItem.Height;
            _oNewItem.Width = _oItem.Width;
            _oNewItem.Left = _oItem.Left + _oItem.Width;

            _oFolderItem = ((SAPbouiCOM.Folder)(_oNewItem.Specific));

            _oFolderItem.Caption = "eBilling";

            _oFolderItem.GroupWith("1320002137");

            ItemsDocuments(_oFormInvoices, _TipoDoc);

            _oFormInvoices.PaneLevel = 1;

        }

        public void AddItemsToBP(SAPbobsCOM.Company _oCompany, SAPbouiCOM.Form oFormBusinessParnerd)
        {
            Funciones.Comunes DLLFunciones = new Funciones.Comunes();

            try
            {
                SAPbouiCOM.Form _oFormBusinessParnerd;
                SAPbouiCOM.Item _oNewItem;
                SAPbouiCOM.Item _oItem;
                SAPbouiCOM.Folder _oFolderItem;

                SAPbobsCOM.Recordset oQuantityEmails = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string sQuantityEmails = null;
                int iQuantityEmails = 0;

                sQuantityEmails = DLLFunciones.GetStringXMLDocument(_oCompany, "eBilling", "eBilling", "GetQuantityEmails");
                oQuantityEmails.DoQuery(sQuantityEmails);

                iQuantityEmails = Convert.ToInt32(oQuantityEmails.Fields.Item(0).Value.ToString());

                _oFormBusinessParnerd = oFormBusinessParnerd;
                _oNewItem = _oFormBusinessParnerd.Items.Add("FolderBO1", SAPbouiCOM.BoFormItemTypes.it_FOLDER);
                _oItem = _oFormBusinessParnerd.Items.Item("234000007");

                _oNewItem.Top = _oItem.Top;
                _oNewItem.Height = _oItem.Height;
                _oNewItem.Width = _oItem.Width;
                _oNewItem.Left = _oItem.Left + _oItem.Width;

                _oFolderItem = ((SAPbouiCOM.Folder)(_oNewItem.Specific));

                _oFolderItem.Caption = "eBilling";

                _oFolderItem.GroupWith("234000007");

                ItemsBusinessParnerd(_oFormBusinessParnerd, iQuantityEmails);

                _oFormBusinessParnerd.PaneLevel = 1;

                DLLFunciones.liberarObjetos(oQuantityEmails);

            }
            catch (Exception e)
            {

                throw;
            }
        }

        private FacturaGeneral oBuillInvoice(SAPbobsCOM.Recordset oCabecera, SAPbobsCOM.Recordset oLineas, SAPbobsCOM.Recordset oImpuestos, SAPbobsCOM.Recordset oImpuestosTotales, SAPbobsCOM.Recordset OCUFEInvoice, string ___TipoDocumento)
        {

            #region Datos Generales Factura

            FacturaGeneral FacturadeVenta = new FacturaGeneral();

            FacturadeVenta.cantidadDecimales = Convert.ToString(oCabecera.Fields.Item("cantidadDecimales").Value.ToString());
            FacturadeVenta.moneda = Convert.ToString(oCabecera.Fields.Item("Moneda").Value.ToString());
            FacturadeVenta.rangoNumeracion = Convert.ToString(oCabecera.Fields.Item("RangoNmeracion").Value.ToString());

            string sredondeoAplicado = Convert.ToString(oCabecera.Fields.Item("RedondeoAplicado").Value.ToString());
            FacturadeVenta.redondeoAplicado = sredondeoAplicado.Replace(",", ".");

            FacturadeVenta.tipoDocumento = Convert.ToString(oCabecera.Fields.Item("tipoDocumento").Value.ToString());
            FacturadeVenta.tipoOperacion = Convert.ToString(oCabecera.Fields.Item("tipoOperacion").Value.ToString());

            string stotalBaseImponible = Convert.ToString(oCabecera.Fields.Item("totalBaseImponible").Value.ToString());
            FacturadeVenta.totalBaseImponible = stotalBaseImponible.Replace(",", ".");

            string stotalBrutoConImpuesto = Convert.ToString(oCabecera.Fields.Item("totalBrutoConImpuesto").Value.ToString());
            FacturadeVenta.totalBrutoConImpuesto = stotalBrutoConImpuesto.Replace(",", ".");

            string stotalMonto = Convert.ToString(oCabecera.Fields.Item("totalMonto").Value.ToString());
            FacturadeVenta.totalMonto = stotalMonto.Replace(",", ".");

            string stotalProductos = Convert.ToString(oCabecera.Fields.Item("totalProductos").Value.ToString());
            FacturadeVenta.totalProductos = stotalProductos.Replace(",", ".");

            string stotalSinImpuestos = Convert.ToString(oCabecera.Fields.Item("totalSinImpuestos").Value.ToString());
            FacturadeVenta.totalSinImpuestos = stotalSinImpuestos.Replace(",", ".");

            #region Descuento generales de la factura

            string ValorDescuentoGeneral = Convert.ToString(oCabecera.Fields.Item("Desc_monto").Value.ToString());
            ValorDescuentoGeneral = ValorDescuentoGeneral.Replace(",", ".");

            if (ValorDescuentoGeneral == "0" || ValorDescuentoGeneral == "0.0" || ValorDescuentoGeneral == "0.00" || ValorDescuentoGeneral == "0.000" || ValorDescuentoGeneral == "0.0000" || ValorDescuentoGeneral == "0.00000" || ValorDescuentoGeneral == "0.000000")
            {

            }
            else
            {
                FacturadeVenta.cargosDescuentos = new CargosDescuentos[1];

                CargosDescuentos DescuentoGeneral = new CargosDescuentos();

                DescuentoGeneral.codigo = Convert.ToString(oCabecera.Fields.Item("Desc_Codigo").Value.ToString());
                DescuentoGeneral.descripcion = Convert.ToString(oCabecera.Fields.Item("Desc_descripcion").Value.ToString());
                DescuentoGeneral.indicador = Convert.ToString(oCabecera.Fields.Item("Desc_indicador").Value.ToString());

                string sDesc_monto = Convert.ToString(oCabecera.Fields.Item("Desc_monto").Value.ToString());
                DescuentoGeneral.monto = sDesc_monto.Replace(",", ".");

                string sDesc_montoBase = Convert.ToString(oCabecera.Fields.Item("Desc_montoBase").Value.ToString());
                DescuentoGeneral.montoBase = sDesc_montoBase.Replace(",", ".");

                string sDesc_porcentaje = Convert.ToString(oCabecera.Fields.Item("Desc_porcentaje").Value.ToString());
                DescuentoGeneral.porcentaje = sDesc_porcentaje.Replace(",", ".");

                DescuentoGeneral.secuencia = Convert.ToString(oCabecera.Fields.Item("Desc_secuencia").Value.ToString());

                FacturadeVenta.cargosDescuentos[0] = DescuentoGeneral;

                FacturadeVenta.totalDescuentos = sDesc_monto.Replace(",", ".");
            }

            #endregion

            FacturadeVenta.consecutivoDocumento = Convert.ToString(oCabecera.Fields.Item("consecutivoDocumento").Value.ToString());
            FacturadeVenta.fechaEmision = oCabecera.Fields.Item("fechaEmision").Value.ToString();

            #endregion

            #region cliente

            Cliente cliente = new Cliente();

            cliente.actividadEconomicaCIIU = Convert.ToString(oCabecera.Fields.Item("actividadEconomicaCIIU").Value.ToString());

            cliente.destinatario = new Destinatario[1];
            Destinatario destinatario1 = new Destinatario();

            destinatario1.canalDeEntrega = Convert.ToString(oCabecera.Fields.Item("canalDeEntrega").Value.ToString());

            #region Revision Correos a Enviar

            #region Variables Correo

            string CorreoDeEntrega1 = Convert.ToString(oCabecera.Fields.Item("correoEntrega1").Value.ToString());
            string CorreoDeEntrega2 = Convert.ToString(oCabecera.Fields.Item("correoEntrega2").Value.ToString());
            string CorreoDeEntrega3 = Convert.ToString(oCabecera.Fields.Item("correoEntrega3").Value.ToString());
            string CorreoDeEntrega4 = Convert.ToString(oCabecera.Fields.Item("correoEntrega4").Value.ToString());
            string CorreoDeEntrega5 = Convert.ToString(oCabecera.Fields.Item("correoEntrega5").Value.ToString());

            int ContadorCorreos = 0;

            #endregion

            #region Contador de los correos a enviar 

            if (string.IsNullOrEmpty(CorreoDeEntrega1))
            {

            }
            else
            {
                ContadorCorreos++;
            }

            if (string.IsNullOrEmpty(CorreoDeEntrega2))
            {

            }
            else
            {
                ContadorCorreos++;
            }

            if (string.IsNullOrEmpty(CorreoDeEntrega3))
            {

            }
            else
            {
                ContadorCorreos++;
            }

            if (string.IsNullOrEmpty(CorreoDeEntrega4))
            {

            }
            else
            {
                ContadorCorreos++;
            }

            if (string.IsNullOrEmpty(CorreoDeEntrega5))
            {

            }
            else
            {
                ContadorCorreos++;
            }

            #endregion

            string[] correoEntrega = new string[ContadorCorreos];

            #region Asignacion de los correos a enviar

            if (ContadorCorreos == 1)
            {
                correoEntrega[0] = CorreoDeEntrega1;
            }
            else if (ContadorCorreos == 2)
            {
                correoEntrega[0] = CorreoDeEntrega1;
                correoEntrega[1] = CorreoDeEntrega2;
            }
            else if (ContadorCorreos == 3)
            {
                correoEntrega[0] = CorreoDeEntrega1;
                correoEntrega[1] = CorreoDeEntrega2;
                correoEntrega[2] = CorreoDeEntrega3;
            }
            else if (ContadorCorreos == 4)
            {
                correoEntrega[0] = CorreoDeEntrega1;
                correoEntrega[1] = CorreoDeEntrega2;
                correoEntrega[2] = CorreoDeEntrega3;
                correoEntrega[3] = CorreoDeEntrega4;
            }
            else if (ContadorCorreos == 5)
            {
                correoEntrega[0] = CorreoDeEntrega1;
                correoEntrega[1] = CorreoDeEntrega2;
                correoEntrega[2] = CorreoDeEntrega3;
                correoEntrega[3] = CorreoDeEntrega4;
                correoEntrega[4] = CorreoDeEntrega5;
            }

            #endregion

            #endregion

            destinatario1.email = correoEntrega;
            destinatario1.fechaProgramada = Convert.ToString(oCabecera.Fields.Item("fechaProgramada").Value.ToString());
            destinatario1.nitProveedorReceptor = Convert.ToString(oCabecera.Fields.Item("nitProveedorReceptor").Value.ToString());
            destinatario1.telefono = Convert.ToString(oCabecera.Fields.Item("telefono").Value.ToString());
            cliente.destinatario[0] = destinatario1;

            cliente.detallesTributarios = new Tributos[1];
            Tributos tributos1 = new Tributos();
            tributos1.codigoImpuesto = Convert.ToString(oCabecera.Fields.Item("codigoImpuesto").Value.ToString());
            cliente.detallesTributarios[0] = tributos1;

            Direccion direccionFiscal = new Direccion();
            direccionFiscal.ciudad = Convert.ToString(oCabecera.Fields.Item("ciudad").Value.ToString());
            direccionFiscal.codigoDepartamento = Convert.ToString(oCabecera.Fields.Item("codigoDepartamento").Value.ToString());
            direccionFiscal.departamento = Convert.ToString(oCabecera.Fields.Item("departamento").Value.ToString());
            direccionFiscal.direccion = Convert.ToString(oCabecera.Fields.Item("direccion").Value.ToString());
            direccionFiscal.lenguaje = Convert.ToString(oCabecera.Fields.Item("lenguaje").Value.ToString());
            direccionFiscal.municipio = Convert.ToString(oCabecera.Fields.Item("municipio").Value.ToString());
            direccionFiscal.pais = Convert.ToString(oCabecera.Fields.Item("pais").Value.ToString());
            direccionFiscal.zonaPostal = Convert.ToString(oCabecera.Fields.Item("zonaPostal").Value.ToString());
            cliente.direccionFiscal = direccionFiscal;
            cliente.direccionCliente = direccionFiscal;

            cliente.email = Convert.ToString(oCabecera.Fields.Item("email").Value.ToString());

            InformacionLegal informacionLegalCliente = new InformacionLegal();
            informacionLegalCliente.codigoEstablecimiento = Convert.ToString(oCabecera.Fields.Item("codigoEstablecimiento").Value.ToString());
            informacionLegalCliente.nombreRegistroRUT = Convert.ToString(oCabecera.Fields.Item("nombreRegistroRUT").Value.ToString());
            informacionLegalCliente.numeroIdentificacion = Convert.ToString(oCabecera.Fields.Item("numeroIdentificacion").Value.ToString());
            informacionLegalCliente.numeroIdentificacionDV = Convert.ToString(oCabecera.Fields.Item("numeroIdentificacionDV").Value.ToString());
            informacionLegalCliente.tipoIdentificacion = Convert.ToString(oCabecera.Fields.Item("tipoIdentificacion").Value.ToString());
            cliente.informacionLegalCliente = informacionLegalCliente;

            cliente.nombreRazonSocial = Convert.ToString(oCabecera.Fields.Item("nombreRazonSocial").Value.ToString());
            cliente.nombreComercial = Convert.ToString(oCabecera.Fields.Item("nombreRazonSocial").Value.ToString());
            cliente.notificar = Convert.ToString(oCabecera.Fields.Item("notificar").Value.ToString());
            cliente.numeroDocumento = Convert.ToString(oCabecera.Fields.Item("numeroDocumento").Value.ToString());
            cliente.numeroIdentificacionDV = Convert.ToString(oCabecera.Fields.Item("numeroIdentificacionDV").Value.ToString());

            cliente.responsabilidadesRut = new Obligaciones[1];
            Obligaciones obligaciones1 = new Obligaciones();
            obligaciones1.obligaciones = Convert.ToString(oCabecera.Fields.Item("obligaciones").Value.ToString());
            obligaciones1.regimen = Convert.ToString(oCabecera.Fields.Item("regimen").Value.ToString());
            cliente.responsabilidadesRut[0] = obligaciones1;

            cliente.tipoIdentificacion = Convert.ToString(oCabecera.Fields.Item("tipoIdentificacion").Value.ToString());
            cliente.tipoPersona = Convert.ToString(oCabecera.Fields.Item("tipoPersona").Value.ToString());

            FacturadeVenta.cliente = cliente;

            #endregion 

            #region Consulta las lineas en la factura de venta

            int CantidadArticulos;
            int SecuenciaArreglo;
            int Posicion;
            CantidadArticulos = oLineas.RecordCount;

            #endregion

            #region Si existen Lineas Asigna los valores de cada columna al arreglo Detalle Factura

            if (CantidadArticulos > 0)
            {
                FacturadeVenta.detalleDeFactura = new FacturaDetalle[CantidadArticulos];

                #region Asignacion Articulos

                oLineas.MoveFirst();

                SecuenciaArreglo = 0;
                Posicion = SecuenciaArreglo + 1;

                do
                {

                    FacturaDetalle Articulo = new FacturaDetalle();

                    #region Detalle articulo

                    Articulo.cantidadPorEmpaque = Convert.ToString(oLineas.Fields.Item("cantidadPorEmpaque").Value.ToString());

                    string scantidadReal = Convert.ToString(oLineas.Fields.Item("cantidadReal").Value.ToString());
                    Articulo.cantidadReal = scantidadReal.Replace(",", ".");

                    Articulo.cantidadRealUnidadMedida = Convert.ToString(oLineas.Fields.Item("cantidadRealUnidadMedida").Value.ToString());

                    string scantidadUnidades = Convert.ToString(oLineas.Fields.Item("cantidadUnidades").Value.ToString());
                    Articulo.cantidadUnidades = scantidadUnidades.Replace(",", ".");

                    Articulo.codigoIdentificadorPais = null;
                    Articulo.codigoProducto = Convert.ToString(oLineas.Fields.Item("codigoProducto").Value.ToString());
                    Articulo.descripcion = Convert.ToString(oLineas.Fields.Item("descripcion").Value.ToString());
                    Articulo.descripcionTecnica = Convert.ToString(oLineas.Fields.Item("descripcion").Value.ToString());
                    Articulo.estandarCodigo = Convert.ToString(oLineas.Fields.Item("estandarCodigo").Value.ToString());
                    Articulo.estandarCodigoProducto = Convert.ToString(oLineas.Fields.Item("estandarCodigoProducto").Value.ToString());

                    #endregion

                    #region Descuentos a nivel de linea

                    string ValorDescuentoLinea = Convert.ToString(oLineas.Fields.Item("Desc_porcentaje").Value);
                    ValorDescuentoLinea = ValorDescuentoLinea.Replace(",", ".");

                    if (ValorDescuentoLinea == "0" || ValorDescuentoLinea == "0.0" || ValorDescuentoLinea == "0.00" || ValorDescuentoLinea == "0.000" || ValorDescuentoLinea == "0.0000" || ValorDescuentoLinea == "0.00000" || ValorDescuentoLinea == "0.000000")
                    {

                    }
                    else
                    {
                        Articulo.cargosDescuentos = new CargosDescuentos[1];

                        CargosDescuentos DescuentoLinea = new CargosDescuentos();

                        DescuentoLinea.descripcion = Convert.ToString(oLineas.Fields.Item("Desc_descripcion").Value.ToString());
                        DescuentoLinea.indicador = Convert.ToString(oLineas.Fields.Item("Desc_indicador").Value.ToString());

                        string sDesc_monto = Convert.ToString(oLineas.Fields.Item("Desc_monto").Value.ToString());
                        DescuentoLinea.monto = sDesc_monto.Replace(",", ".");

                        string sDesc_montoBase = Convert.ToString(oLineas.Fields.Item("Desc_montoBase").Value.ToString());
                        DescuentoLinea.montoBase = sDesc_montoBase.Replace(",", ".");

                        string sPorcentaje = Convert.ToString(oLineas.Fields.Item("Desc_porcentaje").Value.ToString());
                        DescuentoLinea.porcentaje = sPorcentaje.Replace(",", ".");

                        DescuentoLinea.secuencia = Convert.ToString(oLineas.Fields.Item("Desc_secuencia").Value.ToString());

                        Articulo.cargosDescuentos[0] = DescuentoLinea;

                    }

                    #endregion

                    Articulo.impuestosDetalles = new FacturaImpuestos[1];

                    FacturaImpuestos Impuesto = new FacturaImpuestos();

                    #region Detalle Impuesto

                    string sbaseImponibleTOTALImp_Impuesto = Convert.ToString(oLineas.Fields.Item("baseImponibleTOTALImp").Value.ToString());
                    Impuesto.baseImponibleTOTALImp = sbaseImponibleTOTALImp_Impuesto.Replace(",", ".");

                    Impuesto.codigoTOTALImp = Convert.ToString(oLineas.Fields.Item("codigoTOTALImp").Value.ToString());
                    Impuesto.controlInterno = Convert.ToString(oLineas.Fields.Item("controlInterno").Value.ToString());

                    string sporcentajeTOTALImp_Impuesto = Convert.ToString(oLineas.Fields.Item("porcentajeTOTALImp").Value.ToString());
                    Impuesto.porcentajeTOTALImp = sporcentajeTOTALImp_Impuesto.Replace(",", ".");

                    Impuesto.unidadMedida = Convert.ToString(oLineas.Fields.Item("unidadMedida").Value.ToString());
                    Impuesto.unidadMedidaTributo = Convert.ToString(oLineas.Fields.Item("unidadMedidaTributo").Value.ToString());

                    string svalorTOTALImp_Impuesto = Convert.ToString(oLineas.Fields.Item("valorTOTALImp").Value.ToString());
                    Impuesto.valorTOTALImp = svalorTOTALImp_Impuesto.Replace(",", ".");

                    Impuesto.valorTributoUnidad = Convert.ToString(oLineas.Fields.Item("valorTributoUnidad").Value.ToString());

                    #endregion

                    Articulo.impuestosDetalles[0] = Impuesto;

                    Articulo.impuestosTotales = new ImpuestosTotales[1];

                    ImpuestosTotales ImpuestoTOTAL = new ImpuestosTotales();

                    #region Detalle Impuesto Total

                    ImpuestoTOTAL.codigoTOTALImp = Convert.ToString(oLineas.Fields.Item("codigoTOTALImp").Value.ToString());

                    string smontoTotal_ImpuestoTOTAL = Convert.ToString(oLineas.Fields.Item("montoTotal").Value.ToString());
                    ImpuestoTOTAL.montoTotal = smontoTotal_ImpuestoTOTAL.Replace(",", ".");


                    #endregion

                    Articulo.impuestosTotales[0] = ImpuestoTOTAL;

                    #region Demas detalles de la linea del articulo


                    Articulo.marca = Convert.ToString(oLineas.Fields.Item("marca").Value.ToString());
                    Articulo.muestraGratis = Convert.ToString(oLineas.Fields.Item("muestraGratis").Value.ToString());

                    #region Si la linea del articulo es muestra, coloca el tag precio de referencia


                    if (Articulo.muestraGratis == "1")
                    {
                        Articulo.precioReferencia = Convert.ToString(oLineas.Fields.Item("precioReferencia").Value.ToString());
                    }

                    #endregion

                    string sprecioTotal_Articulo = Convert.ToString(oLineas.Fields.Item("precioTotal").Value.ToString());
                    Articulo.precioTotal = sprecioTotal_Articulo.Replace(",", ".");

                    string sprecioTotalSinImpuestos_Articulo = Convert.ToString(oLineas.Fields.Item("precioTotalSinImpuestos").Value.ToString());
                    Articulo.precioTotalSinImpuestos = sprecioTotalSinImpuestos_Articulo.Replace(",", ".");

                    string sprecioVentaUnitario_Articulo = Convert.ToString(oLineas.Fields.Item("precioVentaUnitario").Value.ToString());
                    Articulo.precioVentaUnitario = sprecioVentaUnitario_Articulo.Replace(",", ".");

                    Articulo.secuencia = Convert.ToString(oLineas.Fields.Item("Secuencia").Value.ToString());
                    Articulo.unidadMedida = Convert.ToString(oLineas.Fields.Item("unidadMedida").Value.ToString());


                    #endregion

                    FacturadeVenta.detalleDeFactura[SecuenciaArreglo] = Articulo;

                    SecuenciaArreglo = SecuenciaArreglo + 1;

                    Posicion = Posicion + 1;

                    oLineas.MoveNext();

                } while (oLineas.EoF == false);

                #endregion
            }

            #endregion

            #region Documento Referenciado

            if (___TipoDocumento == "NotaCreditoClientes" || ___TipoDocumento == "NotaDebitoClientes")
            {
                if (Convert.ToString(oCabecera.Fields.Item("tipoOperacion").Value.ToString()) == "22")
                {

                }
                else
                {

                    #region Arreglo donde se asigna comentarios acerca del motivo de la devolucion o anulacion

                    string[] descripcion = new string[1];

                    descripcion[0] = Convert.ToString(oCabecera.Fields.Item("Comentarios_NC").Value.ToString()); ;

                    #endregion

                    FacturadeVenta.documentosReferenciados = new DocumentoReferenciado[2];

                    #region Documento Referenciado 1

                    DocumentoReferenciado Datos_NCoND_DR1 = new DocumentoReferenciado();

                    Datos_NCoND_DR1.codigoEstatusDocumento = Convert.ToString(oCabecera.Fields.Item("codigoEstatusDocumento").Value.ToString());
                    Datos_NCoND_DR1.codigoInterno = "4";
                    Datos_NCoND_DR1.cufeDocReferenciado = Convert.ToString(OCUFEInvoice.Fields.Item("CUFE").Value.ToString());

                    Datos_NCoND_DR1.descripcion = descripcion;

                    Datos_NCoND_DR1.numeroDocumento = Convert.ToString(OCUFEInvoice.Fields.Item("consecutivoDocumento").Value.ToString());

                    FacturadeVenta.documentosReferenciados[0] = Datos_NCoND_DR1;

                    #endregion

                    #region Documento Referenciado 2

                    DocumentoReferenciado Datos_NCoND_DR2 = new DocumentoReferenciado();

                    Datos_NCoND_DR2.codigoInterno = "5";
                    Datos_NCoND_DR2.cufeDocReferenciado = Convert.ToString(OCUFEInvoice.Fields.Item("CUFE").Value.ToString());
                    Datos_NCoND_DR2.fecha = Convert.ToString(OCUFEInvoice.Fields.Item("fechaEmision").Value.ToString());
                    Datos_NCoND_DR2.numeroDocumento = Convert.ToString(OCUFEInvoice.Fields.Item("consecutivoDocumento").Value.ToString());
                    Datos_NCoND_DR2.tipoCUFE = "CUFE-SHA384";

                    FacturadeVenta.documentosReferenciados[1] = Datos_NCoND_DR2;

                    #endregion

                }


            }

            #endregion

            #region Impuestos

            int CantidadImpuestosGenerales;
            int CantidadImpuestosTotales;
            int SecuenciaArregloImpuestos;
            int PosicionImpuestos;

            CantidadImpuestosGenerales = oImpuestos.RecordCount;

            #region Valida si exiten impuestos y asigna a "ImpuestosGenerales"

            if (CantidadImpuestosGenerales > 0)
            {
                FacturadeVenta.impuestosGenerales = new FacturaImpuestos[CantidadImpuestosGenerales];

                oImpuestos.MoveFirst();

                SecuenciaArregloImpuestos = 0;
                PosicionImpuestos = SecuenciaArregloImpuestos + 1;

                do
                {
                    #region Asignacion impuestosGenerales

                    FacturaImpuestos ImpuestosGenerales = new FacturaImpuestos();

                    #region Detalle impuestosGenerales

                    string sbaseImponibleTOTALImp_ImpuestosGenerales = Convert.ToString(oImpuestos.Fields.Item("baseImponibleTOTALImp").Value.ToString());
                    ImpuestosGenerales.baseImponibleTOTALImp = sbaseImponibleTOTALImp_ImpuestosGenerales.Replace(",", ".");

                    ImpuestosGenerales.codigoTOTALImp = Convert.ToString(oImpuestos.Fields.Item("codigoTOTALImp").Value.ToString());

                    string sporcentajeTOTALImp_ImpuestosGenerales = Convert.ToString(oImpuestos.Fields.Item("porcentajeTOTALImp").Value.ToString());
                    ImpuestosGenerales.porcentajeTOTALImp = sporcentajeTOTALImp_ImpuestosGenerales.Replace(",", ".");

                    ImpuestosGenerales.unidadMedida = Convert.ToString(oImpuestos.Fields.Item("unidadMedida").Value.ToString());

                    string svalorTOTALImp_ImpuestosGenerales = Convert.ToString(oImpuestos.Fields.Item("valorTOTALImp").Value.ToString());
                    ImpuestosGenerales.valorTOTALImp = svalorTOTALImp_ImpuestosGenerales.Replace(",", ".");

                    #endregion

                    FacturadeVenta.impuestosGenerales[SecuenciaArregloImpuestos] = ImpuestosGenerales;

                    SecuenciaArregloImpuestos++;
                    PosicionImpuestos++;

                    oImpuestos.MoveNext();

                    #endregion

                } while (oImpuestos.EoF == false);
            }

            #endregion

            #region Valida si exiten impuestos y asigna a "impuestosTotales"

            CantidadImpuestosTotales = oImpuestosTotales.RecordCount;

            if (CantidadImpuestosTotales > 0)
            {
                FacturadeVenta.impuestosTotales = new ImpuestosTotales[CantidadImpuestosTotales];

                oImpuestosTotales.MoveFirst();

                SecuenciaArregloImpuestos = 0;
                PosicionImpuestos = SecuenciaArregloImpuestos + 1;

                do
                {
                    #region Asignacion ImpuestosTotales 

                    ImpuestosTotales ImpuestosTotales = new ImpuestosTotales();

                    #region Detalle ImpuestosTotales

                    ImpuestosTotales.codigoTOTALImp = Convert.ToString(oImpuestosTotales.Fields.Item("codigoTOTALImp").Value.ToString());

                    string smontoTotal_ImpuestosTotales = Convert.ToString(oImpuestosTotales.Fields.Item("valorTOTALImp").Value.ToString());
                    ImpuestosTotales.montoTotal = smontoTotal_ImpuestosTotales.Replace(",", ".");

                    #endregion

                    FacturadeVenta.impuestosTotales[SecuenciaArregloImpuestos] = ImpuestosTotales;

                    SecuenciaArregloImpuestos++;
                    PosicionImpuestos++;

                    oImpuestosTotales.MoveNext();

                    #endregion

                } while (oImpuestosTotales.EoF == false);

            }

            #endregion

            #endregion

            #region mediosDePago

            FacturadeVenta.mediosDePago = new MediosDePago[1];

            MediosDePago MediosPago = new MediosDePago();

            MediosPago.medioPago = Convert.ToString(oCabecera.Fields.Item("medioPago").Value.ToString());
            MediosPago.metodoDePago = Convert.ToString(oCabecera.Fields.Item("FormaPago").Value.ToString());

            if (oCabecera.Fields.Item("FormaPago").Value.ToString() == "2")
            {
                MediosPago.fechaDeVencimiento = Convert.ToString(oCabecera.Fields.Item("fechaDeVencimiento").Value.ToString());
            }

            MediosPago.numeroDeReferencia = Convert.ToString(oCabecera.Fields.Item("numeroDeReferencia").Value.ToString());

            FacturadeVenta.mediosDePago[0] = MediosPago;

            #endregion

            #region Informcion Adicional

            if (Convert.ToString(oCabecera.Fields.Item("FacturaTieneMuestras").Value.ToString()) == "SI")
            {
                string[] txtInformacionAdicional = new string[1];

                txtInformacionAdicional[0] = "El total de la Factura a cobrar corresponde a los items registrado sin considerar la muestra gratis";

                FacturadeVenta.informacionAdicional = txtInformacionAdicional;
            }

            #endregion

            return FacturadeVenta;
        }

        public void ChangueFormPeBilling(SAPbobsCOM.Company oCompany, SAPbouiCOM.Form oFormVDBO, string _sMotor)
        {
            Funciones.Comunes DLLFunciones = new Funciones.Comunes();

            try
            {
                #region Variables y objetos

                SAPbouiCOM.ComboBox _cboTO;

                SAPbouiCOM.Folder oFolder1 = (SAPbouiCOM.Folder)oFormVDBO.Items.Item("Folder1").Specific;
                SAPbouiCOM.PictureBox oLogoBO = (SAPbouiCOM.PictureBox)oFormVDBO.Items.Item("LogoBO").Specific;
                SAPbobsCOM.Recordset oValidValuesTO = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                #endregion

                #region Busqueda de series de numeracion

                SAPbouiCOM.DBDataSource oDBDataSource;

                SAPbouiCOM.Matrix oMatrixSeres = (Matrix)oFormVDBO.Items.Item("MtxSN").Specific;

                oDBDataSource = oFormVDBO.DataSources.DBDataSources.Add("@BOSERNUM");

                oMatrixSeres.Columns.Item("#").Editable = false;

                oMatrixSeres.Columns.Item("Col_01").DataBind.SetBound(true, "@BOSERNUM", "U_BO_TD");
                oMatrixSeres.Columns.Item("Col_02").DataBind.SetBound(true, "@BOSERNUM", "Code");
                oMatrixSeres.Columns.Item("Col_03").DataBind.SetBound(true, "@BOSERNUM", "U_BO_NR");
                oMatrixSeres.Columns.Item("Col_04").DataBind.SetBound(true, "@BOSERNUM", "U_BO_FR");
                oMatrixSeres.Columns.Item("Col_05").DataBind.SetBound(true, "@BOSERNUM", "U_BO_PREF");

                oMatrixSeres.Clear();

                oMatrixSeres.AutoResizeColumns();

                oDBDataSource.Query(null);

                oMatrixSeres.LoadFromDataSource();

                #endregion

                #region Valores validos Tipo de Operacion

                sGetTiposOperacion = DLLFunciones.GetStringXMLDocument(oCompany, "eBilling", "eBilling", "GetTiposOperacion");

                _cboTO = (SAPbouiCOM.ComboBox)oFormVDBO.Items.Item("cboTO").Specific;

                oValidValuesTO.DoQuery(sGetTiposOperacion);
                oValidValuesTO.MoveFirst();

                for (int K = 0; oValidValuesTO.RecordCount - 1 >= K; K++)
                {
                    _cboTO.ValidValues.Add(oValidValuesTO.Fields.Item(0).Value.ToString(), oValidValuesTO.Fields.Item(1).Value.ToString());
                    oValidValuesTO.MoveNext();
                }

                DLLFunciones.liberarObjetos(oValidValuesTO);

                #endregion

                #region Asignacion Logo BO

                oLogoBO.Picture = (Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Core\\Imagenes\\BO.jpg");

                #endregion

                oFormVDBO.DataBrowser.BrowseBy = "txtCode";

                oFolder1.Select();

                oFormVDBO.Visible = true;
                oFormVDBO.Refresh();

            }
            catch (Exception e)
            {
                DLLFunciones.sendErrorMessage(sboapp, e);

            }
        }

        public void ChangueFormUNDM(SAPbouiCOM.Application _sboapp, SAPbobsCOM.Company oCompany, SAPbouiCOM.Form oFormUNDM)
        {
            Funciones.Comunes DLLFunciones = new Funciones.Comunes();

            try
            {

                #region Variables 

                SAPbouiCOM.PictureBox oLogoBO = (SAPbouiCOM.PictureBox)oFormUNDM.Items.Item("LogoBO").Specific;

                #endregion

                #region Asignacion Logo BO

                oLogoBO.Picture = (Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Core\\Imagenes\\BO.jpg");

                #endregion

                #region Centrar el formulario

                oFormUNDM.Left = (_sboapp.Desktop.Width - oFormUNDM.Width) / 2;
                oFormUNDM.Top = (_sboapp.Desktop.Height - oFormUNDM.Height) / 4;

                #endregion

                oFormUNDM.DataBrowser.BrowseBy = "txtUMS";

                oFormUNDM.Visible = true;

                oFormUNDM.Refresh();

            }
            catch (Exception e)
            {
                DLLFunciones.sendErrorMessage(_sboapp, e);
            }
        }

        public void ChangueFormVisoreBilling(SAPbouiCOM.Application _sboapp, SAPbobsCOM.Company oCompany, SAPbouiCOM.Form oFormVDBO, string _sMotor)
        {

            #region Variables y Objetos

            SAPbouiCOM.ComboBox _cboStado = (SAPbouiCOM.ComboBox)oFormVDBO.Items.Item("cboStado").Specific;
            SAPbouiCOM.Folder oFolder1 = (SAPbouiCOM.Folder)oFormVDBO.Items.Item("Folder1").Specific;
            SAPbouiCOM.EditText otxtFI = (SAPbouiCOM.EditText)oFormVDBO.Items.Item("txtFI").Specific;
            SAPbouiCOM.EditText otxtSN = (SAPbouiCOM.EditText)oFormVDBO.Items.Item("txtSN").Specific;

            SAPbouiCOM.PictureBox oLogoBO = (SAPbouiCOM.PictureBox)oFormVDBO.Items.Item("LogoBO").Specific;

            #endregion

            #region Asignacion Logo BO

            oLogoBO.Picture = (Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Core\\Imagenes\\BO.jpg");

            #endregion

            #region Se adicona el ChooFromList 

            oFormVDBO.DataSources.UserDataSources.Add("EditDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);

            AddChooseFromList(_sboapp, oFormVDBO);

            otxtSN.DataBind.SetBound(true, "", "EditDS");
            otxtSN.ChooseFromListUID = "CFL1";
            otxtSN.ChooseFromListAlias = "CardCode";

            #endregion         

            _cboStado.Select("-", BoSearchKey.psk_ByValue);

            oFormVDBO.Left = (_sboapp.Desktop.Width - oFormVDBO.Width) / 2;
            oFormVDBO.Top = (_sboapp.Desktop.Height - oFormVDBO.Height) / 4;

            oFolder1.Select();
            oFormVDBO.Visible = true;

            oFormVDBO.Refresh();

            otxtFI.Item.Click();

        }

        public void LoadFormSendMail(SAPbobsCOM.Company oCompany, SAPbouiCOM.Application _sboapp, SAPbouiCOM.Form oFormSM, string _sMotor, string _sPrefijoDocNumSM)
        {
            Funciones.Comunes DLLFunciones = new Funciones.Comunes();

            try
            {
                #region Variables y objetos

                SAPbouiCOM.PictureBox oLogoBO = (SAPbouiCOM.PictureBox)oFormSM.Items.Item("LogoBO").Specific;

                SAPbouiCOM.Item oFieldSendMail = null;
                SAPbouiCOM.StaticText oStaticText = null;

                SAPbouiCOM.EditText otxtEmail1 = null;

                #endregion

                #region Asignacion Logo BO

                oLogoBO.Picture = (Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Core\\Imagenes\\BO.jpg");

                #endregion

                #region Centrar el formulario

                oFormSM.Left = (_sboapp.Desktop.Width - oFormSM.Width) / 2;
                oFormSM.Top = (_sboapp.Desktop.Height - oFormSM.Height) / 4;

                #endregion

                #region Consultar cantidad de correos 

                string sQuantityEmails;
                int iQuantityEmails;

                SAPbobsCOM.Recordset oQuantityEmails = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                sQuantityEmails = DLLFunciones.GetStringXMLDocument(oCompany, "eBilling", "eBilling", "GetQuantityEmails");
                oQuantityEmails.DoQuery(sQuantityEmails);

                iQuantityEmails = Convert.ToInt32(oQuantityEmails.Fields.Item(0).Value.ToString());

                DLLFunciones.liberarObjetos(oQuantityEmails);

                #endregion

                #region Se obtiene el numero de documento 

                //*******************************************
                // Se adiciona Label con el numero del documento
                //*******************************************

                oFieldSendMail = oFormSM.Items.Add("lbl1", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oFieldSendMail.Left = 10;
                oFieldSendMail.Top = 10;
                oFieldSendMail.Height = 10;
                oFieldSendMail.Width = 10;

                oStaticText = ((SAPbouiCOM.StaticText)(oFieldSendMail.Specific));

                oStaticText.Caption = _sPrefijoDocNumSM;

                oFieldSendMail.Visible = false;

                #endregion

                if (iQuantityEmails == 2)
                {

                    #region Correo 2

                    //*******************************************
                    // Se adiciona Label "Correo Electronico 2"
                    //*******************************************

                    oFieldSendMail = oFormSM.Items.Add("blbEmail2", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    oFieldSendMail.Left = 20;
                    oFieldSendMail.Top = 30;
                    oFieldSendMail.Height = 14;
                    oFieldSendMail.Width = 71;

                    oFieldSendMail.LinkTo = "txtEmail2";

                    oStaticText = ((SAPbouiCOM.StaticText)(oFieldSendMail.Specific));

                    oStaticText.Caption = "Correo 2";

                    //*******************************************
                    // Se adiciona Tex Box "Correo Electronico 2"
                    //*******************************************

                    oFieldSendMail = oFormSM.Items.Add("txtEmail2", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oFieldSendMail.Left = 96;
                    oFieldSendMail.Top = 30;
                    oFieldSendMail.Height = 14;
                    oFieldSendMail.Width = 201;

                    otxtEmail1 = ((SAPbouiCOM.EditText)(oFieldSendMail.Specific));

                    #endregion

                }
                else if (iQuantityEmails == 3)
                {

                    #region Correo 2

                    //*******************************************
                    // Se adiciona Label "Correo Electronico 2"
                    //*******************************************

                    oFieldSendMail = oFormSM.Items.Add("blbEmail2", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    oFieldSendMail.Left = 20;
                    oFieldSendMail.Top = 30;
                    oFieldSendMail.Height = 14;
                    oFieldSendMail.Width = 71;

                    oFieldSendMail.LinkTo = "txtEmail2";

                    oStaticText = ((SAPbouiCOM.StaticText)(oFieldSendMail.Specific));

                    oStaticText.Caption = "Correo 2";

                    //*******************************************
                    // Se adiciona Tex Box "Correo Electronico 1"
                    //*******************************************

                    oFieldSendMail = oFormSM.Items.Add("txtEmail2", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oFieldSendMail.Left = 96;
                    oFieldSendMail.Top = 30;
                    oFieldSendMail.Height = 14;
                    oFieldSendMail.Width = 201;

                    otxtEmail1 = ((SAPbouiCOM.EditText)(oFieldSendMail.Specific));


                    #endregion

                    #region Correo 3

                    //*******************************************
                    // Se adiciona Label "Correo Electronico 3"
                    //*******************************************

                    oFieldSendMail = oFormSM.Items.Add("blbEmail3", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    oFieldSendMail.Left = 20;
                    oFieldSendMail.Top = 45;
                    oFieldSendMail.Height = 14;
                    oFieldSendMail.Width = 71;

                    oFieldSendMail.LinkTo = "txtEmail3";

                    oStaticText = ((SAPbouiCOM.StaticText)(oFieldSendMail.Specific));

                    oStaticText.Caption = "Correo 3";

                    //*******************************************
                    // Se adiciona Tex Box "Correo Electronico 1"
                    //*******************************************

                    oFieldSendMail = oFormSM.Items.Add("txtEmail3", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oFieldSendMail.Left = 96;
                    oFieldSendMail.Top = 45;
                    oFieldSendMail.Height = 14;
                    oFieldSendMail.Width = 201;

                    otxtEmail1 = ((SAPbouiCOM.EditText)(oFieldSendMail.Specific));


                    #endregion

                }
                else if (iQuantityEmails == 4)
                {

                    #region Correo 2

                    //*******************************************
                    // Se adiciona Label "Correo Electronico 2"
                    //*******************************************

                    oFieldSendMail = oFormSM.Items.Add("blbEmail2", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    oFieldSendMail.Left = 20;
                    oFieldSendMail.Top = 30;
                    oFieldSendMail.Height = 14;
                    oFieldSendMail.Width = 71;

                    oFieldSendMail.LinkTo = "txtEmail2";

                    oStaticText = ((SAPbouiCOM.StaticText)(oFieldSendMail.Specific));

                    oStaticText.Caption = "Correo 2";

                    //*******************************************
                    // Se adiciona Tex Box "Correo Electronico 2"
                    //*******************************************

                    oFieldSendMail = oFormSM.Items.Add("txtEmail2", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oFieldSendMail.Left = 96;
                    oFieldSendMail.Top = 30;
                    oFieldSendMail.Height = 14;
                    oFieldSendMail.Width = 201;

                    otxtEmail1 = ((SAPbouiCOM.EditText)(oFieldSendMail.Specific));


                    #endregion

                    #region Correo 3

                    //*******************************************
                    // Se adiciona Label "Correo Electronico 3"
                    //*******************************************

                    oFieldSendMail = oFormSM.Items.Add("blbEmail3", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    oFieldSendMail.Left = 20;
                    oFieldSendMail.Top = 45;
                    oFieldSendMail.Height = 14;
                    oFieldSendMail.Width = 71;

                    oFieldSendMail.LinkTo = "txtEmail3";

                    oStaticText = ((SAPbouiCOM.StaticText)(oFieldSendMail.Specific));

                    oStaticText.Caption = "Correo 3";

                    //*******************************************
                    // Se adiciona Tex Box "Correo Electronico 3"
                    //*******************************************

                    oFieldSendMail = oFormSM.Items.Add("txtEmail3", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oFieldSendMail.Left = 96;
                    oFieldSendMail.Top = 45;
                    oFieldSendMail.Height = 14;
                    oFieldSendMail.Width = 201;

                    otxtEmail1 = ((SAPbouiCOM.EditText)(oFieldSendMail.Specific));


                    #endregion

                    #region Correo 4

                    //*******************************************
                    // Se adiciona Label "Correo Electronico 4"
                    //*******************************************

                    oFieldSendMail = oFormSM.Items.Add("blbEmail4", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    oFieldSendMail.Left = 20;
                    oFieldSendMail.Top = 45;
                    oFieldSendMail.Height = 14;
                    oFieldSendMail.Width = 71;

                    oFieldSendMail.LinkTo = "txtEmail4";

                    oStaticText = ((SAPbouiCOM.StaticText)(oFieldSendMail.Specific));

                    oStaticText.Caption = "Correo 4";

                    //*******************************************
                    // Se adiciona Tex Box "Correo Electronico 4"
                    //*******************************************

                    oFieldSendMail = oFormSM.Items.Add("txtEmail4", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oFieldSendMail.Left = 96;
                    oFieldSendMail.Top = 45;
                    oFieldSendMail.Height = 14;
                    oFieldSendMail.Width = 201;

                    otxtEmail1 = ((SAPbouiCOM.EditText)(oFieldSendMail.Specific));


                    #endregion

                }
                else if (iQuantityEmails == 5)
                {

                    #region Correo 2

                    //*******************************************
                    // Se adiciona Label "Correo Electronico 2"
                    //*******************************************

                    oFieldSendMail = oFormSM.Items.Add("blbEmail2", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    oFieldSendMail.Left = 20;
                    oFieldSendMail.Top = 30;
                    oFieldSendMail.Height = 14;
                    oFieldSendMail.Width = 71;

                    oFieldSendMail.LinkTo = "txtEmail2";

                    oStaticText = ((SAPbouiCOM.StaticText)(oFieldSendMail.Specific));

                    oStaticText.Caption = "Correo 2";

                    //*******************************************
                    // Se adiciona Tex Box "Correo Electronico 2"
                    //*******************************************

                    oFieldSendMail = oFormSM.Items.Add("txtEmail2", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oFieldSendMail.Left = 96;
                    oFieldSendMail.Top = 30;
                    oFieldSendMail.Height = 14;
                    oFieldSendMail.Width = 201;

                    otxtEmail1 = ((SAPbouiCOM.EditText)(oFieldSendMail.Specific));


                    #endregion

                    #region Correo 3

                    //*******************************************
                    // Se adiciona Label "Correo Electronico 3"
                    //*******************************************

                    oFieldSendMail = oFormSM.Items.Add("blbEmail3", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    oFieldSendMail.Left = 20;
                    oFieldSendMail.Top = 45;
                    oFieldSendMail.Height = 14;
                    oFieldSendMail.Width = 71;

                    oFieldSendMail.LinkTo = "txtEmail3";

                    oStaticText = ((SAPbouiCOM.StaticText)(oFieldSendMail.Specific));

                    oStaticText.Caption = "Correo 3";

                    //*******************************************
                    // Se adiciona Tex Box "Correo Electronico 3"
                    //*******************************************

                    oFieldSendMail = oFormSM.Items.Add("txtEmail3", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oFieldSendMail.Left = 96;
                    oFieldSendMail.Top = 45;
                    oFieldSendMail.Height = 14;
                    oFieldSendMail.Width = 201;

                    otxtEmail1 = ((SAPbouiCOM.EditText)(oFieldSendMail.Specific));


                    #endregion

                    #region Correo 4

                    //*******************************************
                    // Se adiciona Label "Correo Electronico 4"
                    //*******************************************

                    oFieldSendMail = oFormSM.Items.Add("blbEmail4", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    oFieldSendMail.Left = 20;
                    oFieldSendMail.Top = 45;
                    oFieldSendMail.Height = 14;
                    oFieldSendMail.Width = 71;

                    oFieldSendMail.LinkTo = "txtEmail4";

                    oStaticText = ((SAPbouiCOM.StaticText)(oFieldSendMail.Specific));

                    oStaticText.Caption = "Correo 4";

                    //*******************************************
                    // Se adiciona Tex Box "Correo Electronico 4"
                    //*******************************************

                    oFieldSendMail = oFormSM.Items.Add("txtEmail4", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oFieldSendMail.Left = 96;
                    oFieldSendMail.Top = 45;
                    oFieldSendMail.Height = 14;
                    oFieldSendMail.Width = 201;

                    otxtEmail1 = ((SAPbouiCOM.EditText)(oFieldSendMail.Specific));


                    #endregion

                    #region Correo 5

                    //*******************************************
                    // Se adiciona Label "Correo Electronico 5"
                    //*******************************************

                    oFieldSendMail = oFormSM.Items.Add("blbEmail5", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    oFieldSendMail.Left = 20;
                    oFieldSendMail.Top = 45;
                    oFieldSendMail.Height = 14;
                    oFieldSendMail.Width = 71;

                    oFieldSendMail.LinkTo = "txtEmail5";

                    oStaticText = ((SAPbouiCOM.StaticText)(oFieldSendMail.Specific));

                    oStaticText.Caption = "Correo 5";

                    //*******************************************
                    // Se adiciona Tex Box "Correo Electronico 5"
                    //*******************************************

                    oFieldSendMail = oFormSM.Items.Add("txtEmail5", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oFieldSendMail.Left = 96;
                    oFieldSendMail.Top = 45;
                    oFieldSendMail.Height = 14;
                    oFieldSendMail.Width = 201;

                    otxtEmail1 = ((SAPbouiCOM.EditText)(oFieldSendMail.Specific));


                    #endregion

                }


                oFormSM.Visible = true;

            }
            catch (Exception e)
            {
                DLLFunciones.sendErrorMessage(sboapp, e);

            }
        }

        public void ChangePaneFolderDocuments(SAPbouiCOM.Form oFormInvoice)
        {
            SAPbouiCOM.Form _oFormInvoice;
            _oFormInvoice = oFormInvoice;
            _oFormInvoice.PaneLevel = 5;
        }

        public void ChangePaneFolderBP(SAPbouiCOM.Form oFormBP)
        {
            SAPbouiCOM.Form _oFormBP;
            _oFormBP = oFormBP;
            _oFormBP.PaneLevel = 28;
        }

        public void CreacionTablasyCamposeBillingBO(SAPbouiCOM.Application sboapp, SAPbobsCOM.Company oCompany, string _sMotor, string _Localizacion)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            try
            {
                int IDFormattedSearchKey = 0;
                string sProcedure_Eliminar = null;
                string sProcedure_Crear = null;
                string sQueryDecimales = null;
                string sCantidadDecimales = null;

                #region Creacion de tablas

                //1
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Tabla - Parametros Iniciales, por favor espere...");
                DllFunciones.crearTabla(oCompany, sboapp, "BOEBILLINGP", "BO-Param. Init. eBilling", SAPbobsCOM.BoUTBTableType.bott_Document);
                //2
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Tabla - Responsabilidades Fiscales, por favor espere...");
                DllFunciones.crearTabla(oCompany, sboapp, "BORESFISCAL", "BO-Responsabilidades Fiscales", SAPbobsCOM.BoUTBTableType.bott_NoObject);
                //3
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Tabla - Unidades de Medida Estandar, por favor espere...");
                DllFunciones.crearTabla(oCompany, sboapp, "BOUNDMED", "BO-Unidades Medida", SAPbobsCOM.BoUTBTableType.bott_MasterData);
                //4
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Tabla - Series de Nuemracion, por favor espere...");
                DllFunciones.crearTabla(oCompany, sboapp, "BOSERNUM", "BO-Series Numeracion", SAPbobsCOM.BoUTBTableType.bott_MasterData);
                //5
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Tabla - Unidades de Medida DIAN, por favor espere...");
                DllFunciones.crearTabla(oCompany, sboapp, "BOUNIDMDIAN", "BO-Unidades de Medida DIAN", SAPbobsCOM.BoUTBTableType.bott_NoObject);
                //6
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Tabla - E-mail Enviados, por favor espere...");
                DllFunciones.crearTabla(oCompany, sboapp, "BOEE", "BO-Email reportados", SAPbobsCOM.BoUTBTableType.bott_NoObject);

                #endregion

                #region Creacion de UDOS

                //6
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando UDO - Parametros Iniciales, por favor espere...");
                string[] TablaseBilling = { "BOEBILLINGP" };
                DllFunciones.CrearUDO(oCompany, sboapp, "BOEBILLINGP", "Parametros iniciales", BoUDOObjType.boud_Document, TablaseBilling, BoYesNoEnum.tNO, BoYesNoEnum.tYES, null, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, 0, 1, BoYesNoEnum.tYES, "BO_eBillingP_Log");
                //7
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando UDO - Unidades Medida, por favor espere...");
                string[] TablaseBilling1 = { "BOUNDMED" };
                DllFunciones.CrearUDO(oCompany, sboapp, "BOUNDMED", "Unidades Medida", BoUDOObjType.boud_MasterData, TablaseBilling1, BoYesNoEnum.tNO, BoYesNoEnum.tYES, null, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, 0, 1, BoYesNoEnum.tYES, "BO_UNDMED_Log");
                //8
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando UDO - Series Numeracion, por favor espere...");
                string[] TablaseBilling2 = { "BOSERNUM" };
                DllFunciones.CrearUDO(oCompany, sboapp, "BOSERNUM", "Series Numeracion", BoUDOObjType.boud_MasterData, TablaseBilling2, BoYesNoEnum.tNO, BoYesNoEnum.tYES, null, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, 0, 1, BoYesNoEnum.tYES, "BO_SERNUM_Log");


                #endregion

                #region Creacion Campos
                //9
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - Unidad Medida DIAN , por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "@BOUNDMED", "BO_UMDIAN", "Token Empresa");
                //10
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - Token Empresa , por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "@BOEBILLINGP", "BO_LlE", "Token Empresa");
                //11
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - Token Password , por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "@BOEBILLINGP", "BO_PwdE", "Token Password");
                //12
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - Modo , por favor espere...");
                string[] ValidValuesFields1 = { "PRO", "PRODUCTIVO", "PRU", "PRUEBAS" };
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 3, "", BoYesNoEnum.tNO, ValidValuesFields1, "@BOEBILLINGP", "BO_Mdo", "Modo");
                //13
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - Activo , por favor espere...");
                string[] ValidValuesFields2 = { "Y", "Si", "N", "No" };
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, "", BoYesNoEnum.tNO, ValidValuesFields2, "@BOEBILLINGP", "BO_Status", "Activo");
                //14
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - Localización utilizada , por favor espere...");
                string[] ValidValuesFields3 = { "OK1", "Consensus", "HBT", "Heinsohn", "EXX", "Exxis", "BO", "Basis One" };
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 3, "", BoYesNoEnum.tNO, ValidValuesFields3, "@BOEBILLINGP", "BO_L", "Localización utilizada");
                //15
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - Forma de Emision , por favor espere...");
                string[] ValidValuesFields11 = { "0", "Sin Adjuntos y R.G. Estandar", "1", "Con Adjuntos y R.G. Estandar", "2", "Con Adjuntos y R.G. Personalizada", "10", "Sin Adjuntos y Sin R.G. Estandar", "11", "Con Adjuntos y sin R.G. Estandar" };
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 2, "", BoYesNoEnum.tNO, ValidValuesFields11, "@BOEBILLINGP", "BO_FormE", "Forma de Emision");
                //16
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - Correo Generico , por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "@BOEBILLINGP", "BO_EmailGen", "Correo Electronico");
                //17
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - Generar XML Prueba , por favor espere...");
                string[] ValidValuesFields20 = { "Y", "Si", "N", "No" };
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, "", BoYesNoEnum.tNO, ValidValuesFields2, "@BOEBILLINGP", "BO_GXP", "Generar XML P.");
                //18
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - Serie Num. Fact , por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "@BOSERNUM", "BO_SN", "Ser. Num. Fac. FE");
                //19
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - Numero de resolución, por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "@BOSERNUM", "BO_NR", "No. Resolución");
                //20
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - Fecha Resolución, por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Date, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "@BOSERNUM", "BO_FR", "Fecha Resolución");
                //21
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - Prefijo, por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "@BOSERNUM", "BO_PREF", "Prefijo");
                //22
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - Tipo de Documento , por favor espere...");
                string[] ValidValuesFields25 = { "FVC", "Factura de venta clientes", "FCC", "Factura de venta contingencia clientes", "NCC", "Nota credito de clientes", "NDC", "Nota debito de clientes" };
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 5, "", BoYesNoEnum.tNO, ValidValuesFields25, "@BOSERNUM", "BO_TD", "Tipo Doc");
                //23
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - Tipo de operacion , por favor espere...");
                string[] ValidValuesFields6 = { "01", "Combustibles", "02", "Emisor es Autorretenedor", "03", "Excluidos y Exentos", "04", "Exportacion", "05", "Generica", "06", "Generica con pago anticipado", "07", "Generica con periodo de facturacion", "08", "Consorcio", "09", "Servicios AIU", "10", "Estandar", "11", "Mandatos bienes", "12", "Mandatos Servicios" };
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 3, "", BoYesNoEnum.tNO, ValidValuesFields6, "@BOEBILLINGP", "BO_TO", "Tipo de Operación");
                //24
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - Prefijo Serie Numeracion , por favor espere...");
                string[] ValidValuesFields12 = { "Y", "Si", "N", "No" };
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 2, "", BoYesNoEnum.tNO, ValidValuesFields12, "@BOEBILLINGP", "BO_Pref", "Pref. Numeracion");
                //25
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - Cantidad Correo , por favor espere...");
                string[] ValidValuesFields7 = { "1", "Hasta 1 Correo", "2", "Hasta 2 Correos", "3", "Hasta 3 Correos", "4", "Hasta 4 Correos", "5", "Hasta 5 Correos", };
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 2, "", BoYesNoEnum.tNO, ValidValuesFields7, "@BOEBILLINGP", "BO_Emails", "Cantidad de Correos");
                //26
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - Resolucion DIAN , por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Date, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "@BOEBILLINGP", "BO_FRDIAN", "Fecha Res. DIAN");
                //27
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - No. Res. DIAN , por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "@BOEBILLINGP", "BO_NRDIAN", "Num. Res. DIAN");
                //28
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - Folios , por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Numeric, BoFldSubTypes.st_None, 10, "", BoYesNoEnum.tNO, null, "@BOEBILLINGP", "BO_Fol", "Folios");
                //29
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - WS Produccion , por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "@BOEBILLINGP", "BO_URLWSPRD", "WEB Services producción");
                //30
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - WS Pruebas , por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "@BOEBILLINGP", "BO_URLWSPRU", "WEB Services pruebas");
                //31
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - Ruta Crystal Report Layout , por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "@BOEBILLINGP", "BO_RutaCRL", "Ruta Crystal Report Layout");
                //32
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - Ruta Crystal Report Informes, por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "@BOEBILLINGP", "BO_RutaCRI", "Ruta Crystal Report Layout");
                //33
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - Ruta XML , por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "@BOEBILLINGP", "BO_RutaXML", "Ruta XML");
                //34
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - Ruta PDF, por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "@BOEBILLINGP", "BO_RutaPDF", "Ruta PDF");
                //35
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - User DB , por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "@BOEBILLINGP", "BO_UserDB", "Usuario DB");
                //36
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - Pass DB , por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "@BOEBILLINGP", "BO_PassDB", "Password DB");
                //37
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - OINV Comentarios , por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "OINV", "BO_EBC", "Comentarios Fac.Elec");
                //38
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - OINV Respuesta , por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 3, "", BoYesNoEnum.tNO, null, "OINV", "BO_CRWS", "Cod. Resp. Fac. Elec");
                //39
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - OINV Mensaje. Res , por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "OINV", "BO_MRWS", "Mens. Resp. Fac. Elec");
                //40
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - OINV PDF Enviado , por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Memo, BoFldSubTypes.st_Link, 100, "", BoYesNoEnum.tNO, null, "OINV", "BO_RPDF", "PDF Enviado");
                //41
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - OINV Estado Doc. , por favor espere...");
                string[] ValidValuesFields4 = { "0", "A la espera", "1", "Aceptada", "2", "Rechazada", "3", "En Validación", "-", "Todos" };
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, "", BoYesNoEnum.tNO, ValidValuesFields4, "OINV", "BO_S", "Estado Documento");
                //42
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - OINV CUFE , por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "OINV", "BO_CUFE", "CUFE");
                //43
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - OINV XML Enviado , por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Memo, BoFldSubTypes.st_Link, 254, "", BoYesNoEnum.tNO, null, "OINV", "BO_XML", "XML Enviado");
                //44
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - OINV Med. Pago , por favor espere...");
                string[] ValidValuesFields9 = { "1", "Instrumento no definido", "2", "Crédito ACH", "3", "Débito ACH", "4", "Reversión débito de demanda ACH", "5", "Reversión crédito de demanda ACH", "6", "Reversión crédito de demanda ACH", "7", "Débito de demanda ACH", "8", "Mantener", "9", "Clearing Nacional o Regional", "10", "Efectivo", "11", "Reversión Crédito Ahorro", "12", "Reversión Débito Ahorro", "13", "Crédito Ahorro", "14", "Débito Ahorro", "15", "Bookentry Crédito", "16", "Bookentry Débito", "17", "Concentración de la demanda en efectivo / Crédito (CCD)", "18", "Concentración de la demanda en efectivo / Debito (CCD)", "19", "Crédito Pago negocio corporativo (CTP)", "20", "Cheque", "21", "Proyecto bancario", "22", "Proyecto bancario certificado", "23", "Cheque bancario", "24", "Nota cambiaria esperando aceptación", "25", "Cheque certificado", "26", "Cheque local", "27", "Débito Pago Negocio Corporativo (CTP)", "28", "Crédito Negocio Intercambio Corporativo (CTX)", "29", "Débito Negocio Intercambio Corporativo (CTX)", "30", "Transferencia Crédito", "31", "Transferencia Débito", "32", "Concentración Efectivo / Desembolso Crédito plus", "33", "Concentración Efectivo / Desembolso Débito plus", "34", "Pago y depósito pre acordado", "35", "Concentración efectivo", "36", "Concentración efectivo ahorros / Desembolso", "37", "Pago Negocio Corporativo Ahorros Crédito", "38", "Pago Negocio Corporativo Ahorros Débito", "39", "Crédito Negocio Intercambio Corporativo", "40", "Débito Negocio Intercambio Corporativo", "41", "Concentración efectivo/Desembolso Crédito plus", "42", "Consignación bancaria", "43", "Concentración efectivo / Desembolso Débito plus", "44", "Nota cambiaria", "45", "Transferencia Crédito Bancario", "46", "Transferencia Débito Interbancario", "47", "Transferencia Débito Bancaria", "48", "Tarjeta Crédito", "49", "Tarjeta Débito", "50", "Pstgiro", "51", "Telex estándar bancario francés", "52", "Pago comercial Urgente", "53", "Pago Tesorería Urgente", "60", "Nota promisoria", "61", "Nota promisoria firmada por el acreedor", "62", "Nota promisoria firmada por el acreedor, avalada por el banco", "63", "Nota promisoria firmada por el acreedor, avalada por un tercero", "64", "Nota promisoria firmada por el banco", "65", "Nota promisoria firmada por un banco, avalada por otro banco", "66", "Nota promisoria firmada", "67", "Nota promisoria firmada por un tercero avalada por un banco", "70", "Retiro de nota por el acreedor", "74", "Retiro de nota por el acreedor sobre un banco", "75", "Retiro de nota por el acreedor, avalada por otro banco", "76", "Retiro de nota por el acreedor, sobre un banco avalada por un tercero", "77", "Retiro de nota por el acreedor sobre un tercero", "78", "Retiro de nota por el acreedor sobre un tercero avalada por un banco", "91", "Nota bancaria transferible", "92", "Cheque local transferible", "93", "Giro referenciado", "94", "Giro Urgente", "95", "Giro formato abierto", "96", "Método de pago solicitado no usado", "97", "Clearing entre partners", "ZZZ", "Acuerdo mutuo" };
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, ValidValuesFields9, "OINV", "BO_MP", "Medio Pago");
                //45
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - OINV Tipo de descuento , por favor espere...");
                string[] ValidValuesFields21 = { "00", "Descuento por impuesto asumido", "01", "Pague uno lleve otro", "02", "Descuentos contractuales", "03", "Descuento por pronto pago", "04", "Envío gratis", "05", "Descuentos específicos por inventarios", "06", "Descuento por monto de compras", "07", "Descuento de temporada", "08", "Descuento por actualización de productos / servicios", "09", "Descuento general", "10", "Descuento por volumen", "11", "Otro descuento" };
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 2, "", BoYesNoEnum.tNO, ValidValuesFields21, "OINV", "BO_DESC", "Tipo Descuento");
                //46
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - OINV Forma de Envio , por favor espere...");
                string[] ValidValuesFields5 = { "A", "AddIn", "M", "Masivo" };
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, "", BoYesNoEnum.tNO, ValidValuesFields5, "OINV", "BO_PP", "Enviado por");
                //47
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - OINV Cod QR , por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Memo, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "OINV", "BO_QR", "Codigo QR");
                //48
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - OINV Enviar E-Mail ? , por favor espere...");
                string[] ValidValuesFields15 = { "Y", "Si", "N", "No" };
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, "", BoYesNoEnum.tNO, ValidValuesFields15, "OINV", "BO_EE", "Enviar E-mail");

                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - OCRD Correo 1 , por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 100, "", BoYesNoEnum.tNO, null, "OCRD", "BO_E_mail_1", "Correo 1");
                //49
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - OCRD Correo 2 , por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 100, "", BoYesNoEnum.tNO, null, "OCRD", "BO_E_mail_2", "Correo 2");
                //50
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - OCRD Correo 3 , por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 100, "", BoYesNoEnum.tNO, null, "OCRD", "BO_E_mail_3", "Correo 3");
                //51
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - OCRD Correo 4 , por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 100, "", BoYesNoEnum.tNO, null, "OCRD", "BO_E_mail_4", "Correo 4");
                //52
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - OCRD Correo 5 , por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 100, "", BoYesNoEnum.tNO, null, "OCRD", "BO_E_mail_5", "Correo 5");
                //53
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - OCRD Resp. Fiscal , por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 10, "", BoYesNoEnum.tNO, null, "OCRD", "BO_RF", "Respon. Fiscal");
                //54
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - OCRD Tip Regimen , por favor espere...");
                string[] ValidValuesFields8 = { "04", "Régimen Simple", "05", "Régimen Ordinario", "48", "Impuesto sobre las ventas - IVA", "49", "No responsable de IVA" };
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 3, "", BoYesNoEnum.tNO, ValidValuesFields8, "OCRD", "BO_TR", "Tipo Regimen");
                //55
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - OCST Codi Departamento , por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 10, "", BoYesNoEnum.tNO, null, "OCST", "BO_CD", "Codigo Departamento");
                //56
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - OINV Descripcion , por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "@BORESFISCAL", "BO_Des", "Descripcion");
                //57
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - ORIN Aplicar a FV No. , por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 50, "", BoYesNoEnum.tNO, null, "ORIN", "BO_AFV", "Aplicar a FV No.");
                //58
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo -  Tipo de nota Credito , por favor espere...");
                if (_Localizacion == "HBT")
                {
                    string[] ValidValuesFields13 = { "1", "Devolucion de Bienes", "2", "Anulación Factura Electronica", "3", "Rebaja Total", "4", "Descuento Total", "5", "Rescisión:", "6", "Otros" };
                    DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 2, "", BoYesNoEnum.tNO, ValidValuesFields13, "ORIN", "BO_TN", "Tipo de Nota");
                }
                //59
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo -  Tipo de nota Debito , por favor espere...");
                if (_Localizacion == "HBT")
                {
                    string[] ValidValuesFields14 = { "1", "Intereses", "2", "Gastos por Cobrar", "3", "Cambio Valor", "4", "Otro" };
                    DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 2, "", BoYesNoEnum.tNO, ValidValuesFields14, "ORIN", "BO_TipND", "Tipo de Nota Debito");
                }

                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - BOEE DocEntry , por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "@BOEE", "BO_DocEntry", "DocEntry");

                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - BOEE ObjecType , por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "@BOEE", "BO_ObjecType", "ObjecType");

                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - BOEE Correo 1 , por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "@BOEE", "BO_Email1", "E-mail 1");

                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - BOEE Correo 2 , por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "@BOEE", "BO_Email2", "E-mail 2");

                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - BOEE Correo 3 , por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "@BOEE", "BO_Email3", "E-mail 3");

                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - BOEE Correo 4 , por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "@BOEE", "BO_Email4", "E-mail 4");

                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Campo - BOEE Correo 5 , por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "@BOEE", "BO_Email5", "E-mail 5");


                #endregion

                #region Creacion Procedures

                //60
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Procedimientos almacenados , por favor espere...");
                SAPbobsCOM.Recordset oProcedures = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                if (_sMotor == "dst_HANADB")
                {
                    #region Consulta si Existente el Procedure

                    sProcedure_Eliminar = DllFunciones.GetStringXMLDocument(oCompany, "eBilling", "Procedures_eBilling", "SearchProcedure");
                    sProcedure_Eliminar = sProcedure_Eliminar.Replace("%sNameProcedure%", "BO_FacturaXML");

                    oProcedures.DoQuery(sProcedure_Eliminar);

                    #endregion

                    if (oProcedures.RecordCount > 0)
                    {
                        #region Elimina el procedure 

                        sProcedure_Eliminar = null;
                        sProcedure_Eliminar = DllFunciones.GetStringXMLDocument(oCompany, "eBilling", "Procedures_eBilling", "Eliminar_BO_FacturaXML");
                        sProcedure_Eliminar = sProcedure_Eliminar.Replace("%sNameProcedure%", "BO_FacturaXML");

                        oProcedures.DoQuery(sProcedure_Eliminar);

                        #endregion

                        #region Crea el procedure

                        if (_Localizacion == "HBT")
                        {
                            sProcedure_Crear = DllFunciones.GetStringXMLDocument(oCompany, "eBilling", "Procedures_eBilling", "BO_FacturaXML_HBT");

                            #region Consulta Decimales

                            sQueryDecimales = DllFunciones.GetStringXMLDocument(oCompany, "eBilling", "eBilling", "DecimalQuantity");
                            oProcedures.DoQuery(sQueryDecimales);

                            sCantidadDecimales = Convert.ToString(oProcedures.Fields.Item("CantidadDecimales").Value.ToString());

                            sProcedure_Crear = sProcedure_Crear.Replace("%Decimal%", sCantidadDecimales);

                            #endregion

                            oProcedures.DoQuery(sProcedure_Crear);

                        }
                        else if (_Localizacion == "OK1")
                        {
                            sProcedure_Crear = DllFunciones.GetStringXMLDocument(oCompany, "eBilling", "Procedures_eBilling", "BO_FacturaXML_OK1");

                            #region Consulta Decimales

                            sQueryDecimales = DllFunciones.GetStringXMLDocument(oCompany, "eBilling", "eBilling", "DecimalQuantity");
                            oProcedures.DoQuery(sQueryDecimales);

                            sCantidadDecimales = Convert.ToString(oProcedures.Fields.Item("CantidadDecimales").Value.ToString());

                            sProcedure_Crear = sProcedure_Crear.Replace("%Decimal%", sCantidadDecimales);

                            #endregion

                            oProcedures.DoQuery(sProcedure_Crear);

                        }
                        else if (_Localizacion == "EXX")
                        {
                            sProcedure_Crear = DllFunciones.GetStringXMLDocument(oCompany, "eBilling", "Procedures_eBilling", "BO_FacturaXML_EXX");
                            oProcedures.DoQuery(sProcedure_Crear);
                        }
                        else if (_Localizacion == "OK1")
                        {
                            sProcedure_Crear = DllFunciones.GetStringXMLDocument(oCompany, "eBilling", "Procedures_eBilling", "BO_FacturaXML_OK1");

                            #region Consulta Decimales

                            sQueryDecimales = DllFunciones.GetStringXMLDocument(oCompany, "eBilling", "eBilling", "DecimalQuantity");
                            oProcedures.DoQuery(sQueryDecimales);

                            sCantidadDecimales = Convert.ToString(oProcedures.Fields.Item("CantidadDecimales").Value.ToString());

                            sProcedure_Crear = sProcedure_Crear.Replace("%Decimal%", sCantidadDecimales);

                            #endregion

                            oProcedures.DoQuery(sProcedure_Crear);

                        }

                        #endregion

                    }
                    else
                    {
                        #region Crea el procedure

                        if (_Localizacion == "HBT")
                        {
                            sProcedure_Crear = DllFunciones.GetStringXMLDocument(oCompany, "eBilling", "Procedures_eBilling", "BO_FacturaXML_HBT");

                            #region Consulta Decimales

                            sQueryDecimales = DllFunciones.GetStringXMLDocument(oCompany, "eBilling", "eBilling", "DecimalQuantity");
                            oProcedures.DoQuery(sQueryDecimales);

                            sCantidadDecimales = Convert.ToString(oProcedures.Fields.Item("CantidadDecimales").Value.ToString());

                            sProcedure_Crear = sProcedure_Crear.Replace("%Decimal%", sCantidadDecimales);

                            #endregion

                            oProcedures.DoQuery(sProcedure_Crear);

                        }
                        else if (_Localizacion == "OK1")
                        {
                            sProcedure_Crear = DllFunciones.GetStringXMLDocument(oCompany, "eBilling", "Procedures_eBilling", "BO_FacturaXML_OK1");
                            oProcedures.DoQuery(sProcedure_Crear);

                        }
                        else if (_Localizacion == "EXX")
                        {
                            sProcedure_Crear = DllFunciones.GetStringXMLDocument(oCompany, "eBilling", "Procedures_eBilling", "BO_FacturaXML_EXX");
                            oProcedures.DoQuery(sProcedure_Crear);
                        }

                        #endregion
                    }
                }
                else
                {
                    #region Consulta si el procedure Existe

                    sProcedure_Eliminar = DllFunciones.GetStringXMLDocument(oCompany, "eBilling", "Procedures_eBilling", "Eliminar_BO_FacturaXML");
                    sProcedure_Eliminar = sProcedure_Eliminar.Replace("%sNameProcedure%", "BO_FacturaXML");

                    oProcedures.DoQuery(sProcedure_Eliminar);

                    #endregion

                    #region Crea el procedure

                    if (_Localizacion == "HBT")
                    {
                        sProcedure_Crear = DllFunciones.GetStringXMLDocument(oCompany, "eBilling", "Procedures_eBilling", "BO_FacturaXML_HBT");
                        oProcedures.DoQuery(sProcedure_Crear);

                    }
                    else if (_Localizacion == "OK1")
                    {
                        sProcedure_Crear = DllFunciones.GetStringXMLDocument(oCompany, "eBilling", "Procedures_eBilling", "BO_FacturaXML_OK1");
                        oProcedures.DoQuery(sProcedure_Crear);

                    }
                    else if (_Localizacion == "EXX")
                    {
                        sProcedure_Crear = DllFunciones.GetStringXMLDocument(oCompany, "eBilling", "Procedures_eBilling", "BO_FacturaXML_EXX");
                        oProcedures.DoQuery(sProcedure_Crear);
                    }
                    else if (_Localizacion == "BO")
                    {
                        sProcedure_Crear = DllFunciones.GetStringXMLDocument(oCompany, "eBilling", "Procedures_eBilling", "BO_FacturaXML_BO");
                        oProcedures.DoQuery(sProcedure_Crear);
                    }

                    #endregion
                }

                DllFunciones.liberarObjetos(oProcedures);
                sProcedure_Crear = null;
                sProcedure_Eliminar = null;

                #endregion

                #region Creacion Busquedas Formateadas
                //61
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Creando Busquedas Formateadas, por favor espere...");

                DllFunciones.AddFormatedSearch(oCompany, sboapp, "eBilling", "Respon. Fiscales", "eBilling", "GetIntrnalKeySearchFormatted", "FormattedSearchResponFiscal");
                DllFunciones.AddFormatedSearch(oCompany, sboapp, "eBilling", "Unid Medida Estandar", "eBilling", "GetIntrnalKeySearchFormatted", "FormattedSearchUnidadesMedidaEstandar");
                DllFunciones.AddFormatedSearch(oCompany, sboapp, "eBilling", "Facturas de Venta", "eBilling", "GetIntrnalKeySearchFormatted", "FormattedSearchFacturasdeVenta");
                DllFunciones.AddFormatedSearch(oCompany, sboapp, "eBilling", "Series Numeracion", "eBilling", "GetIntrnalKeySearchFormatted", "FormattedSearchSeriesNumeracion");
                DllFunciones.AddFormatedSearch(oCompany, sboapp, "eBilling", "Unid Medida DIAN", "eBilling", "GetIntrnalKeySearchFormatted", "FormattedSearchUnidadesMedidaDIANHBT");

                if (_Localizacion == "HBT")
                {
                    DllFunciones.AddFormatedSearch(oCompany, sboapp, "eBilling", "Actividad Economica", "eBilling", "GetIntrnalKeySearchFormatted", "FormattedSearchActividadEconomicaHBT");

                }
                else if (_Localizacion == "OK1")
                {
                    DllFunciones.AddFormatedSearch(oCompany, sboapp, "eBilling", "Actividad Economica", "eBilling", "GetIntrnalKeySearchFormatted", "FormattedSearchActividadEconomicaOK1");
                    DllFunciones.AddFormatedSearch(oCompany, sboapp, "eBilling", "Unid Medida DIAN", "eBilling", "GetIntrnalKeySearchFormatted", "FormattedSearchUnidadesMedidaDIANOK1");
                }
                else if (_Localizacion == "EXX")
                {
                    DllFunciones.AddFormatedSearch(oCompany, sboapp, "eBilling", "Actividad Economica", "eBilling", "GetIntrnalKeySearchFormatted", "FormattedSearchActividadEconomicaEXX");
                    DllFunciones.AddFormatedSearch(oCompany, sboapp, "eBilling", "Unid Medida DIAN", "eBilling", "GetIntrnalKeySearchFormatted", "FormattedSearchUnidadesMedidaDIANEXX");
                }
                else if (_Localizacion == "BO")
                {
                    DllFunciones.AddFormatedSearch(oCompany, sboapp, "eBilling", "Actividad Economica", "eBilling", "GetIntrnalKeySearchFormatted", "FormattedSearchActividadEconomicaBO");
                    DllFunciones.AddFormatedSearch(oCompany, sboapp, "eBilling", "Unid Medida DIAN", "eBilling", "GetIntrnalKeySearchFormatted", "FormattedSearchUnidadesMedidaDIANBO");
                }
                #endregion

                #region Asignacion Busquedas Formateadas
                //62
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Asignando Busquedas Formateadas, por favor espere...");

                #region Actividad Economica tabla parametros eBilling

                IDFormattedSearchKey = 0;

                IDFormattedSearchKey = DllFunciones.GetFormmatedSearchKey("UDO_FT_BO_eBillingP", "txtAC", oCompany, sboapp);

                if (IDFormattedSearchKey == 0)
                {
                    #region Se adiciona la busqueda formateada al campo 

                    SAPbobsCOM.FormattedSearches oSFActividadEconomica = (SAPbobsCOM.FormattedSearches)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oFormattedSearches);
                    SAPbobsCOM.Recordset oFormattedSearched = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    oSFActividadEconomica.FormID = "UDO_FT_BO_eBillingP";
                    oSFActividadEconomica.ItemID = "txtAC";
                    oSFActividadEconomica.Action = SAPbobsCOM.BoFormattedSearchActionEnum.bofsaQuery;
                    oSFActividadEconomica.FieldID = "txtAC";
                    oSFActividadEconomica.ColumnID = "-1";

                    sGetFormattedSearch = DllFunciones.GetStringXMLDocument(oCompany, "eBilling", "eBilling", "GetIntrnalKeySearchFormatted");
                    _IDCategory = DllFunciones.SearchCatetoryID(oCompany, "eBilling", "eBilling");
                    sGetFormattedSearch = sGetFormattedSearch.Replace("%CategoryID%", _IDCategory).Replace("%NameSearchFormatted%", "Actividad Economica");

                    oFormattedSearched.DoQuery(sGetFormattedSearch);

                    oSFActividadEconomica.QueryID = Convert.ToInt32(oFormattedSearched.Fields.Item(0).Value.ToString());

                    oSFActividadEconomica.Refresh = SAPbobsCOM.BoYesNoEnum.tYES;
                    oSFActividadEconomica.ForceRefresh = SAPbobsCOM.BoYesNoEnum.tYES;
                    oSFActividadEconomica.ByField = SAPbobsCOM.BoYesNoEnum.tYES;

                    Rsd = oSFActividadEconomica.Add();

                    if (Rsd == 0)
                    {
                        DllFunciones.liberarObjetos(oFormattedSearched);
                    }
                    else
                    {
                        DllFunciones.sendMessageBox(sboapp, oCompany.GetLastErrorDescription());
                        DllFunciones.liberarObjetos(oFormattedSearched);
                    }

                    #endregion
                }
                else
                {

                }

                #endregion

                #region Series Numeracion Matrix parametros eBilling

                IDFormattedSearchKey = 0;

                IDFormattedSearchKey = DllFunciones.GetFormmatedSearchKey("UDO_FT_BO_eBillingP", "MtxSN", oCompany, sboapp);

                if (IDFormattedSearchKey == 0)
                {
                    #region Se adiciona la busqueda formateada al campo 

                    SAPbobsCOM.FormattedSearches oSFActividadEconomica = (SAPbobsCOM.FormattedSearches)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oFormattedSearches);
                    SAPbobsCOM.Recordset oFormattedSearched = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    oSFActividadEconomica.FormID = "UDO_FT_BO_eBillingP";
                    oSFActividadEconomica.ItemID = "MtxSN";
                    oSFActividadEconomica.Action = SAPbobsCOM.BoFormattedSearchActionEnum.bofsaQuery;
                    oSFActividadEconomica.FieldID = "MtxSN";
                    oSFActividadEconomica.ColumnID = "Col_02";

                    sGetFormattedSearch = DllFunciones.GetStringXMLDocument(oCompany, "eBilling", "eBilling", "GetIntrnalKeySearchFormatted");
                    _IDCategory = DllFunciones.SearchCatetoryID(oCompany, "eBilling", "eBilling");
                    sGetFormattedSearch = sGetFormattedSearch.Replace("%CategoryID%", _IDCategory).Replace("%NameSearchFormatted%", "Series Numeracion");

                    oFormattedSearched.DoQuery(sGetFormattedSearch);

                    oSFActividadEconomica.QueryID = Convert.ToInt32(oFormattedSearched.Fields.Item(0).Value.ToString());

                    oSFActividadEconomica.Refresh = SAPbobsCOM.BoYesNoEnum.tYES;
                    oSFActividadEconomica.ForceRefresh = SAPbobsCOM.BoYesNoEnum.tYES;
                    oSFActividadEconomica.ByField = SAPbobsCOM.BoYesNoEnum.tYES;

                    Rsd = oSFActividadEconomica.Add();

                    if (Rsd == 0)
                    {
                        DllFunciones.liberarObjetos(oFormattedSearched);
                    }
                    else
                    {
                        DllFunciones.sendMessageBox(sboapp, oCompany.GetLastErrorDescription());
                        DllFunciones.liberarObjetos(oFormattedSearched);
                    }

                    #endregion
                }
                else
                {

                }

                #endregion

                #region Responsabilidades Fiscales 

                IDFormattedSearchKey = 0;

                IDFormattedSearchKey = DllFunciones.GetFormmatedSearchKey("134", "txtRF", oCompany, sboapp);

                if (IDFormattedSearchKey == 0)
                {
                    #region Se adiciona la busqueda formateada al campo 

                    SAPbobsCOM.FormattedSearches oSFActividadEconomica = (SAPbobsCOM.FormattedSearches)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oFormattedSearches);
                    SAPbobsCOM.Recordset oFormattedSearched = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    oSFActividadEconomica.FormID = "134";
                    oSFActividadEconomica.ItemID = "txtRF";
                    oSFActividadEconomica.Action = SAPbobsCOM.BoFormattedSearchActionEnum.bofsaQuery;
                    oSFActividadEconomica.FieldID = "txtRF";
                    oSFActividadEconomica.ColumnID = "-1";

                    sGetFormattedSearch = DllFunciones.GetStringXMLDocument(oCompany, "eBilling", "eBilling", "GetIntrnalKeySearchFormatted");
                    _IDCategory = DllFunciones.SearchCatetoryID(oCompany, "eBilling", "eBilling");
                    sGetFormattedSearch = sGetFormattedSearch.Replace("%CategoryID%", _IDCategory).Replace("%NameSearchFormatted%", "Respon. Fiscales");

                    oFormattedSearched.DoQuery(sGetFormattedSearch);

                    oSFActividadEconomica.QueryID = Convert.ToInt32(oFormattedSearched.Fields.Item(0).Value.ToString());

                    oSFActividadEconomica.Refresh = SAPbobsCOM.BoYesNoEnum.tYES;
                    oSFActividadEconomica.ForceRefresh = SAPbobsCOM.BoYesNoEnum.tYES;
                    oSFActividadEconomica.ByField = SAPbobsCOM.BoYesNoEnum.tYES;

                    Rsd = oSFActividadEconomica.Add();

                    if (Rsd == 0)
                    {
                        DllFunciones.liberarObjetos(oFormattedSearched);
                    }
                    else
                    {
                        DllFunciones.sendMessageBox(sboapp, oCompany.GetLastErrorDescription());
                        DllFunciones.liberarObjetos(oFormattedSearched);
                    }

                    #endregion
                }
                else
                {
                }

                #endregion

                #region Codigo Unidad de Medida Estandar 

                IDFormattedSearchKey = 0;

                IDFormattedSearchKey = DllFunciones.GetFormmatedSearchKey("UDO_FT_BOUNDMED", "txtUMS", oCompany, sboapp);

                if (IDFormattedSearchKey == 0)
                {
                    #region Se adiciona la busqueda formateada al campo 

                    SAPbobsCOM.FormattedSearches oSFActividadEconomica = (SAPbobsCOM.FormattedSearches)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oFormattedSearches);
                    SAPbobsCOM.Recordset oFormattedSearched = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    oSFActividadEconomica.FormID = "UDO_FT_BOUNDMED";
                    oSFActividadEconomica.ItemID = "txtUMS";
                    oSFActividadEconomica.Action = SAPbobsCOM.BoFormattedSearchActionEnum.bofsaQuery;
                    oSFActividadEconomica.FieldID = "txtUMS";
                    oSFActividadEconomica.ColumnID = "-1";

                    sGetFormattedSearch = DllFunciones.GetStringXMLDocument(oCompany, "eBilling", "eBilling", "GetIntrnalKeySearchFormatted");
                    _IDCategory = DllFunciones.SearchCatetoryID(oCompany, "eBilling", "eBilling");
                    sGetFormattedSearch = sGetFormattedSearch.Replace("%CategoryID%", _IDCategory).Replace("%NameSearchFormatted%", "Unid Medida Estandar");

                    oFormattedSearched.DoQuery(sGetFormattedSearch);

                    oSFActividadEconomica.QueryID = Convert.ToInt32(oFormattedSearched.Fields.Item(0).Value.ToString());

                    oSFActividadEconomica.Refresh = SAPbobsCOM.BoYesNoEnum.tYES;
                    oSFActividadEconomica.ForceRefresh = SAPbobsCOM.BoYesNoEnum.tYES;
                    oSFActividadEconomica.ByField = SAPbobsCOM.BoYesNoEnum.tYES;

                    Rsd = oSFActividadEconomica.Add();

                    if (Rsd == 0)
                    {
                        DllFunciones.liberarObjetos(oFormattedSearched);
                    }
                    else
                    {
                        DllFunciones.sendMessageBox(sboapp, oCompany.GetLastErrorDescription());
                        DllFunciones.liberarObjetos(oFormattedSearched);
                    }

                    #endregion
                }
                else
                {
                }

                #endregion

                #region Codigo Unidad de Medida DIAN 

                IDFormattedSearchKey = 0;

                IDFormattedSearchKey = DllFunciones.GetFormmatedSearchKey("UDO_FT_BOUNDMED", "txtUMD", oCompany, sboapp);

                if (IDFormattedSearchKey == 0)
                {
                    #region Se adiciona la busqueda formateada al campo 

                    SAPbobsCOM.FormattedSearches oSFActividadEconomica = (SAPbobsCOM.FormattedSearches)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oFormattedSearches);
                    SAPbobsCOM.Recordset oFormattedSearched = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    oSFActividadEconomica.FormID = "UDO_FT_BOUNDMED";
                    oSFActividadEconomica.ItemID = "txtUMD";
                    oSFActividadEconomica.Action = SAPbobsCOM.BoFormattedSearchActionEnum.bofsaQuery;
                    oSFActividadEconomica.FieldID = "txtUMD";
                    oSFActividadEconomica.ColumnID = "-1";

                    sGetFormattedSearch = DllFunciones.GetStringXMLDocument(oCompany, "eBilling", "eBilling", "GetIntrnalKeySearchFormatted");
                    _IDCategory = DllFunciones.SearchCatetoryID(oCompany, "eBilling", "eBilling");
                    sGetFormattedSearch = sGetFormattedSearch.Replace("%CategoryID%", _IDCategory).Replace("%NameSearchFormatted%", "Unid Medida DIAN");

                    oFormattedSearched.DoQuery(sGetFormattedSearch);

                    oSFActividadEconomica.QueryID = Convert.ToInt32(oFormattedSearched.Fields.Item(0).Value.ToString());

                    oSFActividadEconomica.Refresh = SAPbobsCOM.BoYesNoEnum.tYES;
                    oSFActividadEconomica.ForceRefresh = SAPbobsCOM.BoYesNoEnum.tYES;
                    oSFActividadEconomica.ByField = SAPbobsCOM.BoYesNoEnum.tYES;

                    Rsd = oSFActividadEconomica.Add();

                    if (Rsd == 0)
                    {
                        DllFunciones.liberarObjetos(oFormattedSearched);
                    }
                    else
                    {
                        DllFunciones.sendMessageBox(sboapp, oCompany.GetLastErrorDescription());
                        DllFunciones.liberarObjetos(oFormattedSearched);
                    }

                    #endregion
                }
                else
                {

                }

                #endregion

                #region Facturas de venta - Nota Credito

                IDFormattedSearchKey = 0;

                IDFormattedSearchKey = DllFunciones.GetFormmatedSearchKey("179", "txtAFV", oCompany, sboapp);

                if (IDFormattedSearchKey == 0)
                {
                    #region Se adiciona la busqueda formateada al campo 

                    SAPbobsCOM.FormattedSearches oSFActividadEconomica = (SAPbobsCOM.FormattedSearches)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oFormattedSearches);
                    SAPbobsCOM.Recordset oFormattedSearched = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    oSFActividadEconomica.FormID = "179";
                    oSFActividadEconomica.ItemID = "txtAFV";
                    oSFActividadEconomica.Action = SAPbobsCOM.BoFormattedSearchActionEnum.bofsaQuery;
                    oSFActividadEconomica.FieldID = "txtAFV";
                    oSFActividadEconomica.ColumnID = "-1";

                    sGetFormattedSearch = DllFunciones.GetStringXMLDocument(oCompany, "eBilling", "eBilling", "GetIntrnalKeySearchFormatted");
                    _IDCategory = DllFunciones.SearchCatetoryID(oCompany, "eBilling", "eBilling");
                    sGetFormattedSearch = sGetFormattedSearch.Replace("%CategoryID%", _IDCategory).Replace("%NameSearchFormatted%", "Facturas de Venta");

                    oFormattedSearched.DoQuery(sGetFormattedSearch);

                    oSFActividadEconomica.QueryID = Convert.ToInt32(oFormattedSearched.Fields.Item(0).Value.ToString());

                    oSFActividadEconomica.Refresh = SAPbobsCOM.BoYesNoEnum.tYES;
                    oSFActividadEconomica.ForceRefresh = SAPbobsCOM.BoYesNoEnum.tYES;
                    oSFActividadEconomica.ByField = SAPbobsCOM.BoYesNoEnum.tYES;

                    Rsd = oSFActividadEconomica.Add();

                    if (Rsd == 0)
                    {
                        DllFunciones.liberarObjetos(oFormattedSearched);
                    }
                    else
                    {
                        DllFunciones.sendMessageBox(sboapp, oCompany.GetLastErrorDescription());
                        DllFunciones.liberarObjetos(oFormattedSearched);
                    }

                    #endregion
                }
                else
                {
                }

                #endregion

                #endregion

                #region Importacion Archivos CSV
                //63
                DllFunciones.ProgressBar(oCompany, sboapp, 72, 1, "Importando archivos CSV, por favor espere...");

                DllFunciones.ImportCSV(sboapp, oCompany, "Tiposresponsabilidades", "eBilling", "GetTableBORESFISCAL", "InsertTipoResponsabilidad");

                DllFunciones.ImportCSV(sboapp, oCompany, "UnidadesdeMedidaDIAN", "eBilling", "GetTableUMDIAN", "InsertUnidadMedidaDIAN");

                #endregion

            }
            catch (Exception e)
            {
                DllFunciones.sendErrorMessage(sboapp, e);
            }

        }

        public void ConsultaTokens(SAPbobsCOM.Company _oCompany, SAPbouiCOM.Application sboapp, SAPbouiCOM.Form _oFormParametros)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            try
            {
                #region Consulta URL

                string sGetModo = null;
                string sURLEmision = null;
                string sURLAdjuntos = null;
                string sModo = null;
                string sDocEntry = null;

                SAPbobsCOM.Recordset oConsultarGetModo = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                sDocEntry = ((SAPbouiCOM.EditText)(_oFormParametros.Items.Item("txtCode").Specific)).Value.ToString();

                sDocEntry = sDocEntry.Trim();

                sGetModo = DllFunciones.GetStringXMLDocument(_oCompany, "eBilling", "eBilling", "GetModoandURL");

                sGetModo = sGetModo.Replace("%Estado%", " ").Replace("%DocEntry%", "\"DocEntry\" = '" + sDocEntry + "'");

                oConsultarGetModo.DoQuery(sGetModo);

                sURLEmision = Convert.ToString(oConsultarGetModo.Fields.Item("URLTFHKA").Value.ToString()) + "/ws/v1.0/Service.svc?wsdl";
                sURLAdjuntos = Convert.ToString(oConsultarGetModo.Fields.Item("URLTFHKA").Value.ToString()) + "/ws/adjuntos/Service.svc?wsdl";
                sModo = Convert.ToString(oConsultarGetModo.Fields.Item("Modo").Value.ToString());

                DllFunciones.liberarObjetos(oConsultarGetModo);

                #endregion

                #region Instanciacion parametros TFHKA

                //Especifica el puerto (HTTP o HTTPS)
                if (sModo == "PRU")
                {
                    BasicHttpBinding port = new BasicHttpBinding();
                }
                else if (sModo == "PRO")
                {
                    BasicHttpsBinding port = new BasicHttpsBinding();
                }

                port.MaxBufferPoolSize = Int32.MaxValue;
                port.MaxBufferSize = Int32.MaxValue;
                port.MaxReceivedMessageSize = Int32.MaxValue;
                port.ReaderQuotas.MaxStringContentLength = Int32.MaxValue;
                port.SendTimeout = TimeSpan.FromMinutes(2);
                port.ReceiveTimeout = TimeSpan.FromMinutes(2);

                if (sModo == "PRO")
                {
                    port.Security.Mode = BasicHttpSecurityMode.Transport;
                }

                //Especifica la dirección de conexion para Demo y Adjuntos para pruebas
                EndpointAddress endPointEmision = new EndpointAddress(sURLEmision); //URL DEMO EMISION

                ServicioEmisionFE.ServiceClient serviceClienTFHKA;

                serviceClienTFHKA = new ServicioEmisionFE.ServiceClient(port, endPointEmision);

                #endregion

                #region Variables y Objetos

                SAPbouiCOM.EditText otxtFol;
                SAPbouiCOM.EditText otxtLlE;
                SAPbouiCOM.EditText otxtPwdE;

                otxtFol = (EditText)_oFormParametros.Items.Item("txtFol").Specific;
                otxtLlE = (EditText)_oFormParametros.Items.Item("txtLlE").Specific;
                otxtPwdE = (EditText)_oFormParametros.Items.Item("txtPwdE").Specific;

                #endregion

                FoliosRemainingResponse Tokens = serviceClienTFHKA.FoliosRestantes(otxtLlE.Value, otxtPwdE.Value);

                if (Tokens.codigo == 200)
                {
                    #region Actualiza los campos en el formulario

                    otxtFol.Value = Convert.ToString(Tokens.foliosRestantes);
                    otxtFol.Item.Enabled = false;
                    _oFormParametros.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    _oFormParametros.Refresh();
                    DllFunciones.sendMessageBox(sboapp, "Tokens sincronizados con TFHKA correctamente");

                    #endregion

                }
                else
                {
                    DllFunciones.sendMessageBox(sboapp, Tokens.mensaje);
                }

            }
            catch (Exception e)
            {

                DllFunciones.sendMessageBox(sboapp, e.Message);

            }
        }

        public Boolean ExportPDF(SAPbouiCOM.Application _sboapp, SAPbobsCOM.Company _oCompany, string _RutaPDFyXML, string _DocEntry, string _RutaCR, string __TipoDocumento, string _sUserDB, string _sPassDB)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            try
            {
                #region Variables  y objetos

                string sGetRPTDoc = null;
                string sGetRPTDocUser = null;
                string sRutaLayout = null;
                string _sMotorDB = null;
                string _sServer = null;
                string _sNameDB = null;
                string _sTipo = null;
                string _UserId = null;
                string _strConnection = null;
                string _sArquitectura = null;

                if (__TipoDocumento == "FacturaDeClientes")
                {
                    _sTipo = "INV2";
                }
                else if (__TipoDocumento == "NotaCreditoClientes")
                {
                    _sTipo = "RIN2";
                }
                else if (__TipoDocumento == "NotaDebitoClientes")
                {
                    _sTipo = "IDN2";
                }

                SAPbobsCOM.Recordset oRGetRPTDoc = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                #endregion

                #region Consulta del Motor de Base de datos Y Nombre Base de datos y arquitectura

                _sMotorDB = Convert.ToString(_oCompany.DbServerType);
                _sNameDB = Convert.ToString(_oCompany.CompanyDB);
                _sServer = Convert.ToString(_oCompany.Server);
                _sArquitectura = Convert.ToString(System.IntPtr.Size);

                #endregion

                #region Consulta del nombre del Formato RPT y la ruta donde se encuentra ubicado el RPT

                _UserId = Convert.ToString(_oCompany.UserSignature);

                sGetRPTDoc = DllFunciones.GetStringXMLDocument(_oCompany, "eBilling", "eBilling", "GetRPTDocUser");
                sGetRPTDoc = sGetRPTDoc.Replace("%TypeDoc%", _sTipo).Replace("%UserId%", _UserId);

                oRGetRPTDoc.DoQuery(sGetRPTDoc);

                if (oRGetRPTDoc.RecordCount > 0)
                {

                }
                else
                {
                    sGetRPTDoc = DllFunciones.GetStringXMLDocument(_oCompany, "eBilling", "eBilling", "GetRPTDoc");
                    sGetRPTDoc = sGetRPTDoc.Replace("%TypeDoc%", _sTipo);

                    oRGetRPTDoc.DoQuery(sGetRPTDoc);
                }

                sRutaLayout = _RutaCR + "\\" + Convert.ToString(oRGetRPTDoc.Fields.Item("NombreFormato").Value.ToString()) + ".rpt";

                #endregion

                #region Generacion del PDF

                if (_sMotorDB == "dst_HANADB")
                {
                    if (_sArquitectura == "8")
                    {
                        #region Genera el PDF con cliente SAP a 64X

                        ReportDocument LayoutPDF = new ReportDocument();

                        LayoutPDF.Load(sRutaLayout);

                        _strConnection = string.Format("DRIVER={0};UID={1};PWD={2};SERVERNODE={3};DATABASE={4};", "{B1CRHPROXY}", _sUserDB, _sPassDB, _sServer, _sNameDB);

                        NameValuePairs2 logonProps2 = LayoutPDF.DataSourceConnections[0].LogonProperties;
                        logonProps2.Set("Provider", "B1CRHPROXY");
                        logonProps2.Set("Server Type", "B1CRHPROXY");
                        logonProps2.Set("Connection String", _strConnection);

                        LayoutPDF.DataSourceConnections[0].SetLogonProperties(logonProps2);
                        LayoutPDF.DataSourceConnections[0].SetConnection(_sServer, _sNameDB, false);
                        LayoutPDF.SetParameterValue("DocKey@", _DocEntry);
                        LayoutPDF.SetParameterValue("Schema@", _sNameDB);

                        LayoutPDF.ExportToDisk(ExportFormatType.PortableDocFormat, _RutaPDFyXML);

                        LayoutPDF.Close();

                        LayoutPDF.Dispose();

                        GC.SuppressFinalize(LayoutPDF);

                        #endregion
                    }
                    else if (_sArquitectura == "4")
                    {
                        #region Genera el PDF con cliente SAP a 32X

                        ReportDocument LayoutPDF = new ReportDocument();

                        LayoutPDF.Load(sRutaLayout);

                        _strConnection = string.Format("DRIVER={0};UID={1};PWD={2};SERVERNODE={3};DATABASE={4};", "{B1CRHPROXY32}", _sUserDB, _sPassDB, _sServer, _sNameDB);

                        NameValuePairs2 logonProps2 = LayoutPDF.DataSourceConnections[0].LogonProperties;
                        logonProps2.Set("Provider", "B1CRHPROXY32");
                        logonProps2.Set("Server Type", "B1CRHPROXY32");
                        logonProps2.Set("Connection String", _strConnection);

                        LayoutPDF.DataSourceConnections[0].SetLogonProperties(logonProps2);
                        LayoutPDF.DataSourceConnections[0].SetConnection(_sServer, _sNameDB, false);
                        LayoutPDF.SetParameterValue("DocKey@", _DocEntry);
                        LayoutPDF.SetParameterValue("Schema@", _sNameDB);

                        LayoutPDF.ExportToDisk(ExportFormatType.PortableDocFormat, _RutaPDFyXML);

                        LayoutPDF.Close();

                        LayoutPDF.Dispose();

                        GC.SuppressFinalize(LayoutPDF);

                        #endregion
                    }
                }
                else
                {
                    #region Genera el PDF con cliente SAP a 32x o 64x

                    ReportDocument LayoutPDF = new ReportDocument();

                    DiskFileDestinationOptions DestinoDocumento = new DiskFileDestinationOptions();
                    PdfRtfWordFormatOptions OpcionesPDF = new PdfRtfWordFormatOptions();

                    LayoutPDF.Load(sRutaLayout);

                    int Contador = LayoutPDF.DataSourceConnections.Count;
                    LayoutPDF.DataSourceConnections[0].IntegratedSecurity = false;
                    LayoutPDF.DataSourceConnections[0].SetLogon(_sUserDB, _sPassDB);
                    ExportOptions OpExport = LayoutPDF.ExportOptions;
                    OpExport.ExportDestinationType = ExportDestinationType.DiskFile;
                    OpExport.ExportFormatType = ExportFormatType.PortableDocFormat;
                    DestinoDocumento.DiskFileName = _RutaPDFyXML;
                    OpExport.ExportDestinationOptions = (ExportDestinationOptions)DestinoDocumento;
                    OpExport.ExportFormatOptions = (ExportFormatOptions)OpcionesPDF;

                    LayoutPDF.SetParameterValue("DocKey@", _DocEntry);

                    LayoutPDF.Export();

                    LayoutPDF.Close();

                    LayoutPDF.Dispose();

                    GC.SuppressFinalize(LayoutPDF);

                    #endregion
                }

                #endregion

                #region Libreacion de Objetos

                DllFunciones.liberarObjetos(oRGetRPTDoc);

                #endregion

            }
            catch (Exception e)
            {
                DllFunciones.sendErrorMessage(_sboapp, e);
            }

            return true;
        }

        public void InsertDataInMatrix(SAPbouiCOM.Application _sboapp, SAPbobsCOM.Company oCompany, SAPbouiCOM.Form oFormMatrixInovice)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            #region Varibles y objetos

            SAPbouiCOM.ComboBox _cboStado = (SAPbouiCOM.ComboBox)oFormMatrixInovice.Items.Item("cboStado").Specific;

            SAPbouiCOM.EditText oFI = (SAPbouiCOM.EditText)oFormMatrixInovice.Items.Item("txtFI").Specific;
            SAPbouiCOM.EditText oFF = (SAPbouiCOM.EditText)oFormMatrixInovice.Items.Item("txtFF").Specific;
            SAPbouiCOM.EditText oDocNum = (SAPbouiCOM.EditText)oFormMatrixInovice.Items.Item("txtND").Specific;
            SAPbouiCOM.EditText oSN = (SAPbouiCOM.EditText)oFormMatrixInovice.Items.Item("txtSN").Specific;

            SAPbouiCOM.Button oBtnVD = (SAPbouiCOM.Button)oFormMatrixInovice.Items.Item("btnVD").Specific;

            #endregion

            if (string.IsNullOrEmpty(oFI.Value))
            {
                DllFunciones.sendMessageBox(_sboapp, "Por favor seleccionar la Fecha Inicial");
            }
            else if (string.IsNullOrEmpty(oFF.Value))
            {
                DllFunciones.sendMessageBox(_sboapp, "Por favor seleccionar la Fecha Final");
            }
            else
            {

                #region Parametros Generales

                string sPath;
                string sInvoices = null;
                string sCreditMemo = null;
                string sDebitMemo = null;
                string sSeriesNumber = null;
                string sQuantityEmails = null;
                int iCount;
                string EstadoDocsaConsultar = null;
                int CantidadRegistos = 0;

                sPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

                sSeriesNumber = DllFunciones.GetStringXMLDocument(oCompany, "eBilling", "eBilling", "GetSeriesNumberActive");

                sQuantityEmails = DllFunciones.GetStringXMLDocument(oCompany, "eBilling", "eBilling", "QuantityEmails");

                #endregion

                #region Consulta series de numeración configuradas como Facturacion Electronica

                SAPbobsCOM.Recordset oRecorsetSeriesNumbers = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset oQuantityEmails = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                oRecorsetSeriesNumbers.DoQuery(sSeriesNumber);
                oQuantityEmails.DoQuery(sQuantityEmails);

                #endregion

                #region Consulta de facturas, notas debito y notas credito a mostrar en matrix

                SAPbouiCOM.Matrix oMatrixInvoice = (Matrix)oFormMatrixInovice.Items.Item("MtxOINV").Specific;
                SAPbouiCOM.Matrix oMatrixCreditMemo = (Matrix)oFormMatrixInovice.Items.Item("MtxORIN").Specific;
                SAPbouiCOM.Matrix oMatrixDebitMemo = (Matrix)oFormMatrixInovice.Items.Item("MtxOINVD").Specific;

                SAPbobsCOM.Recordset oRecorsetInvoices = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset oRecorsetCreditMemo = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset oRecorsetDebitMemo = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                SAPbouiCOM.DataTable oTableInvoices = oFormMatrixInovice.DataSources.DataTables.Item("DT_Invoices");
                SAPbouiCOM.DataTable oTableCreditMemo = oFormMatrixInovice.DataSources.DataTables.Item("DT_CreditMemo");
                SAPbouiCOM.DataTable oTableDebitMemo = oFormMatrixInovice.DataSources.DataTables.Item("DT_DebitMemo");

                sInvoices = DllFunciones.GetStringXMLDocument(oCompany, "eBilling", "eBilling", "GetInvoices");
                sCreditMemo = DllFunciones.GetStringXMLDocument(oCompany, "eBilling", "eBilling", "GetCreditMemo");
                sDebitMemo = DllFunciones.GetStringXMLDocument(oCompany, "eBilling", "eBilling", "GetDebitMemo");

                EstadoDocsaConsultar = _cboStado.Value.ToString();

                #region Pasa parametros Fecha Incial, Fecha Fnal y Estado Documento

                if (EstadoDocsaConsultar == "-")
                {
                    sInvoices = sInvoices.Replace("%FI%", oFI.Value.ToString()).Replace("%FF%", oFF.Value.ToString()).Replace("%EstadoDocumento%", " ").Replace("%Series%", Convert.ToString(oRecorsetSeriesNumbers.Fields.Item(0).Value));
                    sCreditMemo = sCreditMemo.Replace("%FI%", oFI.Value.ToString()).Replace("%FF%", oFF.Value.ToString()).Replace("%EstadoDocumento%", " ").Replace("%Series%", Convert.ToString(oRecorsetSeriesNumbers.Fields.Item(1).Value));
                    sDebitMemo = sDebitMemo.Replace("%FI%", oFI.Value.ToString()).Replace("%FF%", oFF.Value.ToString()).Replace("%EstadoDocumento%", " ").Replace("%Series%", Convert.ToString(oRecorsetSeriesNumbers.Fields.Item(2).Value));
                }
                else
                {
                    sInvoices = sInvoices.Replace("%FI%", oFI.Value.ToString()).Replace("%FF%", oFF.Value.ToString()).Replace("%EstadoDocumento%", "AND \"U_BO_S\" IN ('" + _cboStado.Value.ToString() + "')").Replace("%=%", "=").Replace("%Series%", Convert.ToString(oRecorsetSeriesNumbers.Fields.Item(0).Value));
                    sCreditMemo = sCreditMemo.Replace("%FI%", oFI.Value.ToString()).Replace("%FF%", oFF.Value.ToString()).Replace("%EstadoDocumento%", "AND \"U_BO_S\" IN ('" + _cboStado.Value.ToString() + "')").Replace("%=%", "=").Replace("%Series%", Convert.ToString(oRecorsetSeriesNumbers.Fields.Item(1).Value));
                    sDebitMemo = sDebitMemo.Replace("%FI%", oFI.Value.ToString()).Replace("%FF%", oFF.Value.ToString()).Replace("%EstadoDocumento%", "AND \"U_BO_S\" IN ('" + _cboStado.Value.ToString() + "')").Replace("%Series%", Convert.ToString(oRecorsetSeriesNumbers.Fields.Item(2).Value));
                }
                #endregion

                #region Pasa parametro Socio Negocio

                if (string.IsNullOrWhiteSpace(oSN.Value))
                {
                    sInvoices = sInvoices.Replace("***SN***", "");
                    sCreditMemo = sCreditMemo.Replace("***SN***", "");
                    sDebitMemo = sDebitMemo.Replace("***SN***", "");

                }
                else
                {
                    sInvoices = sInvoices.Replace("***SN***", "AND \"CardCode\" = '" + oSN.Value + "'");
                    sCreditMemo = sCreditMemo.Replace("***SN***", "AND \"CardCode\" = '" + oSN.Value + "'");
                    sDebitMemo = sDebitMemo.Replace("***SN***", "AND \"CardCode\" = '" + oSN.Value + "'");
                }

                #endregion

                #region pasa parametro Numero Documento

                if (string.IsNullOrWhiteSpace(oDocNum.Value))
                {
                    sInvoices = sInvoices.Replace("***DocNum***", "");
                    sCreditMemo = sCreditMemo.Replace("***DocNum***", "");
                    sDebitMemo = sDebitMemo.Replace("***DocNum***", "");

                }
                else
                {
                    sInvoices = sInvoices.Replace("***DocNum***", "AND \"DocNum\" = '" + oDocNum.Value + "'");
                    sCreditMemo = sCreditMemo.Replace("***DocNum***", "AND \"DocNum\" = '" + oDocNum.Value + "'");
                    sDebitMemo = sDebitMemo.Replace("***DocNum***", "AND \"DocNum\" = '" + oDocNum.Value + "'");
                }

                #endregion

                #region pasa parametro cantidad de correos a mostrar

                iCount = Convert.ToInt32(oQuantityEmails.Fields.Item("CantidadCorreos").Value.ToString());

                #endregion

                oRecorsetInvoices.DoQuery(sInvoices);

                oRecorsetCreditMemo.DoQuery(sCreditMemo);

                oRecorsetDebitMemo.DoQuery(sDebitMemo);

                oTableInvoices.ExecuteQuery(sInvoices);
                oTableCreditMemo.ExecuteQuery(sCreditMemo);
                oTableDebitMemo.ExecuteQuery(sDebitMemo);

                CantidadRegistos = oRecorsetInvoices.RecordCount + oRecorsetCreditMemo.RecordCount + oRecorsetDebitMemo.RecordCount;

                #endregion

                if (CantidadRegistos != 0)
                {
                    #region Carga datos Matrix Facturas

                    if (oRecorsetInvoices.RecordCount > 0)
                    {
                        oMatrixInvoice.Clear();

                        oMatrixInvoice.Columns.Item("#").DataBind.Bind("DT_Invoices", "#");
                        oMatrixInvoice.Columns.Item("Col_0").DataBind.Bind("DT_Invoices", "Estado");
                        oMatrixInvoice.Columns.Item("Col_9").DataBind.Bind("DT_Invoices", "DocEntry");
                        oMatrixInvoice.Columns.Item("Col_1").DataBind.Bind("DT_Invoices", "No_Factura");
                        oMatrixInvoice.Columns.Item("Col_2").DataBind.Bind("DT_Invoices", "Codigo_cliente");
                        oMatrixInvoice.Columns.Item("Col_3").DataBind.Bind("DT_Invoices", "Nombre_cliente");
                        oMatrixInvoice.Columns.Item("Col_4").DataBind.Bind("DT_Invoices", "Fecha_Documento");
                        oMatrixInvoice.Columns.Item("Col_5").DataBind.Bind("DT_Invoices", "Fecha_vencimiento");
                        oMatrixInvoice.Columns.Item("Col_5").DataBind.Bind("DT_Invoices", "Fecha_vencimiento");
                        oMatrixInvoice.Columns.Item("Col_10").DataBind.Bind("DT_Invoices", "Email_Enviado");

                        if (iCount == 1)
                        {
                            oMatrixInvoice.Columns.Item("Col_11").DataBind.Bind("DT_Invoices", "Correo1");
                            oMatrixInvoice.Columns.Item("Col_12").Visible = false;
                            oMatrixInvoice.Columns.Item("Col_13").Visible = false;
                            oMatrixInvoice.Columns.Item("Col_14").Visible = false;
                            oMatrixInvoice.Columns.Item("Col_15").Visible = false;
                        }
                        else if (iCount == 2)
                        {
                            oMatrixInvoice.Columns.Item("Col_11").DataBind.Bind("DT_Invoices", "Correo1");
                            oMatrixInvoice.Columns.Item("Col_12").DataBind.Bind("DT_Invoices", "Correo2");
                            oMatrixInvoice.Columns.Item("Col_12").Visible = true;
                            oMatrixInvoice.Columns.Item("Col_13").Visible = false;
                            oMatrixInvoice.Columns.Item("Col_14").Visible = false;
                            oMatrixInvoice.Columns.Item("Col_15").Visible = false;
                        }
                        else if (iCount == 3)
                        {
                            oMatrixInvoice.Columns.Item("Col_11").DataBind.Bind("DT_Invoices", "Correo1");
                            oMatrixInvoice.Columns.Item("Col_12").DataBind.Bind("DT_Invoices", "Correo2");
                            oMatrixInvoice.Columns.Item("Col_12").Visible = true;
                            oMatrixInvoice.Columns.Item("Col_13").DataBind.Bind("DT_Invoices", "Correo3");
                            oMatrixInvoice.Columns.Item("Col_13").Visible = true;
                            oMatrixInvoice.Columns.Item("Col_14").Visible = false;
                            oMatrixInvoice.Columns.Item("Col_15").Visible = false;
                        }
                        else if (iCount == 4)
                        {
                            oMatrixInvoice.Columns.Item("Col_11").DataBind.Bind("DT_Invoices", "Correo1");
                            oMatrixInvoice.Columns.Item("Col_12").DataBind.Bind("DT_Invoices", "Correo2");
                            oMatrixInvoice.Columns.Item("Col_12").Visible = true;
                            oMatrixInvoice.Columns.Item("Col_13").DataBind.Bind("DT_Invoices", "Correo3");
                            oMatrixInvoice.Columns.Item("Col_13").Visible = true;
                            oMatrixInvoice.Columns.Item("Col_14").DataBind.Bind("DT_Invoices", "Correo4");
                            oMatrixInvoice.Columns.Item("Col_14").Visible = true;
                            oMatrixInvoice.Columns.Item("Col_15").Visible = false;
                        }
                        else if (iCount == 5)
                        {
                            oMatrixInvoice.Columns.Item("Col_11").DataBind.Bind("DT_Invoices", "Correo1");
                            oMatrixInvoice.Columns.Item("Col_12").DataBind.Bind("DT_Invoices", "Correo2");
                            oMatrixInvoice.Columns.Item("Col_12").Visible = true;
                            oMatrixInvoice.Columns.Item("Col_13").DataBind.Bind("DT_Invoices", "Correo3");
                            oMatrixInvoice.Columns.Item("Col_13").Visible = true;
                            oMatrixInvoice.Columns.Item("Col_14").DataBind.Bind("DT_Invoices", "Correo4");
                            oMatrixInvoice.Columns.Item("Col_14").Visible = true;
                            oMatrixInvoice.Columns.Item("Col_15").DataBind.Bind("DT_Invoices", "Correo5");
                            oMatrixInvoice.Columns.Item("Col_15").Visible = true;
                        }

                        oMatrixInvoice.Columns.Item("Col_6").DataBind.Bind("DT_Invoices", "Total_documento");
                        oMatrixInvoice.Columns.Item("Col_7").DataBind.Bind("DT_Invoices", "Respuesta_TFHKA");

                        oMatrixInvoice.LoadFromDataSource();

                        oMatrixInvoice.AutoResizeColumns();

                    }

                    #endregion

                    #region Carga datos Matrix Notas credito

                    if (oRecorsetCreditMemo.RecordCount > 0)
                    {
                        oMatrixCreditMemo.Clear();

                        oMatrixCreditMemo.Columns.Item("#").DataBind.Bind("DT_CreditMemo", "#");
                        oMatrixCreditMemo.Columns.Item("Col_0").DataBind.Bind("DT_CreditMemo", "Estado");
                        oMatrixCreditMemo.Columns.Item("Col_9").DataBind.Bind("DT_CreditMemo", "DocEntry");
                        oMatrixCreditMemo.Columns.Item("Col_1").DataBind.Bind("DT_CreditMemo", "No_Factura");
                        oMatrixCreditMemo.Columns.Item("Col_2").DataBind.Bind("DT_CreditMemo", "Codigo_cliente");
                        oMatrixCreditMemo.Columns.Item("Col_3").DataBind.Bind("DT_CreditMemo", "Nombre_cliente");
                        oMatrixCreditMemo.Columns.Item("Col_4").DataBind.Bind("DT_CreditMemo", "Fecha_Documento");
                        oMatrixCreditMemo.Columns.Item("Col_5").DataBind.Bind("DT_CreditMemo", "Fecha_vencimiento");
                        oMatrixCreditMemo.Columns.Item("Col_10").DataBind.Bind("DT_CreditMemo", "Email_Enviado");

                        if (iCount == 1)
                        {
                            oMatrixCreditMemo.Columns.Item("Col_11").DataBind.Bind("DT_CreditMemo", "Correo1");
                            oMatrixCreditMemo.Columns.Item("Col_12").Visible = false;
                            oMatrixCreditMemo.Columns.Item("Col_13").Visible = false;
                            oMatrixCreditMemo.Columns.Item("Col_14").Visible = false;
                            oMatrixCreditMemo.Columns.Item("Col_15").Visible = false;
                        }
                        else if (iCount == 2)
                        {
                            oMatrixCreditMemo.Columns.Item("Col_11").DataBind.Bind("DT_CreditMemo", "Correo1");
                            oMatrixCreditMemo.Columns.Item("Col_12").DataBind.Bind("DT_CreditMemo", "Correo2");
                            oMatrixCreditMemo.Columns.Item("Col_12").Visible = true;
                            oMatrixCreditMemo.Columns.Item("Col_13").Visible = false;
                            oMatrixCreditMemo.Columns.Item("Col_14").Visible = false;
                            oMatrixCreditMemo.Columns.Item("Col_15").Visible = false;
                        }
                        else if (iCount == 3)
                        {
                            oMatrixCreditMemo.Columns.Item("Col_11").DataBind.Bind("DT_CreditMemo", "Correo1");
                            oMatrixCreditMemo.Columns.Item("Col_12").DataBind.Bind("DT_CreditMemo", "Correo2");
                            oMatrixCreditMemo.Columns.Item("Col_12").Visible = true;
                            oMatrixCreditMemo.Columns.Item("Col_13").DataBind.Bind("DT_CreditMemo", "Correo3");
                            oMatrixCreditMemo.Columns.Item("Col_13").Visible = true;
                            oMatrixCreditMemo.Columns.Item("Col_14").Visible = false;
                            oMatrixCreditMemo.Columns.Item("Col_15").Visible = false;
                        }
                        else if (iCount == 4)
                        {
                            oMatrixCreditMemo.Columns.Item("Col_11").DataBind.Bind("DT_CreditMemo", "Correo1");
                            oMatrixCreditMemo.Columns.Item("Col_12").DataBind.Bind("DT_CreditMemo", "Correo2");
                            oMatrixCreditMemo.Columns.Item("Col_12").Visible = true;
                            oMatrixCreditMemo.Columns.Item("Col_13").DataBind.Bind("DT_CreditMemo", "Correo3");
                            oMatrixCreditMemo.Columns.Item("Col_13").Visible = true;
                            oMatrixCreditMemo.Columns.Item("Col_14").DataBind.Bind("DT_CreditMemo", "Correo4");
                            oMatrixCreditMemo.Columns.Item("Col_14").Visible = true;
                            oMatrixCreditMemo.Columns.Item("Col_15").Visible = false;
                        }
                        else if (iCount == 5)
                        {
                            oMatrixCreditMemo.Columns.Item("Col_11").DataBind.Bind("DT_CreditMemo", "Correo1");
                            oMatrixCreditMemo.Columns.Item("Col_12").DataBind.Bind("DT_CreditMemo", "Correo2");
                            oMatrixCreditMemo.Columns.Item("Col_12").Visible = true;
                            oMatrixCreditMemo.Columns.Item("Col_13").DataBind.Bind("DT_CreditMemo", "Correo3");
                            oMatrixCreditMemo.Columns.Item("Col_13").Visible = true;
                            oMatrixCreditMemo.Columns.Item("Col_14").DataBind.Bind("DT_CreditMemo", "Correo4");
                            oMatrixCreditMemo.Columns.Item("Col_14").Visible = true;
                            oMatrixCreditMemo.Columns.Item("Col_15").DataBind.Bind("DT_CreditMemo", "Correo5");
                            oMatrixCreditMemo.Columns.Item("Col_15").Visible = true;
                        }

                        oMatrixCreditMemo.Columns.Item("Col_6").DataBind.Bind("DT_CreditMemo", "Total_documento");
                        oMatrixCreditMemo.Columns.Item("Col_7").DataBind.Bind("DT_CreditMemo", "Respuesta_TFHKA");

                        oMatrixCreditMemo.LoadFromDataSource();

                        oMatrixCreditMemo.AutoResizeColumns();

                    }
                    #endregion

                    #region Carga datos Matrix Notas Debito

                    if (oRecorsetDebitMemo.RecordCount > 0)
                    {
                        oMatrixDebitMemo.Clear();

                        oMatrixDebitMemo.Columns.Item("#").DataBind.Bind("DT_DebitMemo", "#");
                        oMatrixDebitMemo.Columns.Item("Col_0").DataBind.Bind("DT_DebitMemo", "Estado");
                        oMatrixDebitMemo.Columns.Item("Col_9").DataBind.Bind("DT_DebitMemo", "DocEntry");
                        oMatrixDebitMemo.Columns.Item("Col_1").DataBind.Bind("DT_DebitMemo", "No_Factura");
                        oMatrixDebitMemo.Columns.Item("Col_2").DataBind.Bind("DT_DebitMemo", "Codigo_cliente");
                        oMatrixDebitMemo.Columns.Item("Col_3").DataBind.Bind("DT_DebitMemo", "Nombre_cliente");
                        oMatrixDebitMemo.Columns.Item("Col_4").DataBind.Bind("DT_DebitMemo", "Fecha_Documento");
                        oMatrixDebitMemo.Columns.Item("Col_5").DataBind.Bind("DT_DebitMemo", "Fecha_vencimiento");

                        oMatrixDebitMemo.Columns.Item("Col_5").DataBind.Bind("DT_DebitMemo", "Email_Enviado");

                        if (iCount == 1)
                        {
                            oMatrixDebitMemo.Columns.Item("Col_11").DataBind.Bind("DT_DebitMemo", "Correo1");
                            oMatrixDebitMemo.Columns.Item("Col_12").Visible = false;
                            oMatrixDebitMemo.Columns.Item("Col_13").Visible = false;
                            oMatrixDebitMemo.Columns.Item("Col_14").Visible = false;
                            oMatrixDebitMemo.Columns.Item("Col_15").Visible = false;
                        }
                        else if (iCount == 2)
                        {
                            oMatrixDebitMemo.Columns.Item("Col_11").DataBind.Bind("DT_DebitMemo", "Correo1");
                            oMatrixDebitMemo.Columns.Item("Col_12").DataBind.Bind("DT_DebitMemo", "Correo2");
                            oMatrixDebitMemo.Columns.Item("Col_13").Visible = false;
                            oMatrixDebitMemo.Columns.Item("Col_14").Visible = false;
                            oMatrixDebitMemo.Columns.Item("Col_15").Visible = false;

                        }
                        else if (iCount == 3)
                        {
                            oMatrixDebitMemo.Columns.Item("Col_11").DataBind.Bind("DT_DebitMemo", "Correo1");
                            oMatrixDebitMemo.Columns.Item("Col_12").DataBind.Bind("DT_DebitMemo", "Correo2");
                            oMatrixDebitMemo.Columns.Item("Col_13").DataBind.Bind("DT_DebitMemo", "Correo3");
                            oMatrixDebitMemo.Columns.Item("Col_14").Visible = false;
                            oMatrixDebitMemo.Columns.Item("Col_15").Visible = false;
                        }
                        else if (iCount == 4)
                        {
                            oMatrixDebitMemo.Columns.Item("Col_11").DataBind.Bind("DT_DebitMemo", "Correo1");
                            oMatrixDebitMemo.Columns.Item("Col_12").DataBind.Bind("DT_DebitMemo", "Correo2");
                            oMatrixDebitMemo.Columns.Item("Col_13").DataBind.Bind("DT_DebitMemo", "Correo3");
                            oMatrixDebitMemo.Columns.Item("Col_14").DataBind.Bind("DT_DebitMemo", "Correo4");
                            oMatrixDebitMemo.Columns.Item("Col_15").Visible = false;
                        }
                        else if (iCount == 5)
                        {
                            oMatrixDebitMemo.Columns.Item("Col_11").DataBind.Bind("DT_DebitMemo", "Correo1");
                            oMatrixDebitMemo.Columns.Item("Col_12").DataBind.Bind("DT_DebitMemo", "Correo2");
                            oMatrixDebitMemo.Columns.Item("Col_13").DataBind.Bind("DT_DebitMemo", "Correo3");
                            oMatrixDebitMemo.Columns.Item("Col_14").DataBind.Bind("DT_DebitMemo", "Correo4");
                            oMatrixDebitMemo.Columns.Item("Col_15").DataBind.Bind("DT_DebitMemo", "Correo5");
                        }

                        oMatrixDebitMemo.Columns.Item("Col_6").DataBind.Bind("DT_DebitMemo", "Total_documento");
                        oMatrixDebitMemo.Columns.Item("Col_7").DataBind.Bind("DT_DebitMemo", "Respuesta_TFHKA");

                        oMatrixDebitMemo.LoadFromDataSource();

                        oMatrixDebitMemo.AutoResizeColumns();

                    }
                    #endregion

                    oBtnVD.Item.Enabled = true;
                }
                else
                {
                    DllFunciones.sendMessageBox(_sboapp, "No se encontraron documentos");
                }

                #region Liberacion de Objetos

                DllFunciones.liberarObjetos(oRecorsetInvoices);
                DllFunciones.liberarObjetos(oRecorsetCreditMemo);
                DllFunciones.liberarObjetos(oRecorsetDebitMemo);
                DllFunciones.liberarObjetos(oRecorsetSeriesNumbers);
                DllFunciones.liberarObjetos(oQuantityEmails);

                #endregion

            }

            oFormMatrixInovice.Refresh();
        }

        public Boolean Insert_InfoUDO_eBillingP(SAPbouiCOM.Application sboapp, SAPbobsCOM.Company oCompany, SAPbouiCOM.Form oFormParametros, string _sMotor)
        {

            try
            {
                Funciones.Comunes DllFunciones = new Funciones.Comunes();

                #region Variables y objetos

                EditText oDocEntry;
                EditText otxtLlE;
                EditText otxtPwdE;
                EditText otxtDocEntry;
                SAPbouiCOM.ComboBox otxtMdo;
                SAPbouiCOM.ComboBox otxtL;
                SAPbouiCOM.Matrix oMatrixSeres;
                SAPbouiCOM.CheckBox oChkStatus;

                SAPbobsCOM.Recordset oGetLastRecord = oGetLastRecord = ((SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));
                SAPbobsCOM.Recordset oActiveConfig = oGetLastRecord = ((SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));

                oDocEntry = (EditText)oFormParametros.Items.Item("txtCode").Specific;
                otxtLlE = (EditText)oFormParametros.Items.Item("txtLlE").Specific;
                otxtPwdE = (EditText)oFormParametros.Items.Item("txtPwdE").Specific;
                otxtDocEntry = (EditText)oFormParametros.Items.Item("txtCode").Specific;
                otxtMdo = (SAPbouiCOM.ComboBox)oFormParametros.Items.Item("txtMdo").Specific;
                otxtL = (SAPbouiCOM.ComboBox)oFormParametros.Items.Item("txtL").Specific;
                oMatrixSeres = (Matrix)oFormParametros.Items.Item("MtxSN").Specific;
                oChkStatus = (SAPbouiCOM.CheckBox)oFormParametros.Items.Item("txtStatus").Specific;

                string sActiveConfig = null;
                string sDocEntryRecorset = null;
                int iCounterActiveConfig = 0;

                #endregion

                #region Consultar configuraciones activas

                sActiveConfig = DllFunciones.GetStringXMLDocument(oCompany, "eBilling", "eBilling", "GetActiveConfig");

                oActiveConfig.DoQuery(sActiveConfig);

                sDocEntryRecorset = Convert.ToString(oActiveConfig.Fields.Item("DocEntry").Value.ToString());

                if (sDocEntryRecorset != otxtDocEntry.Value.ToString())
                {
                    if (oChkStatus.Checked == true)
                    {
                        iCounterActiveConfig = oActiveConfig.RecordCount + 1;
                    }
                }

                #endregion

                #region Validación de campos obligatorios

                if (string.IsNullOrEmpty(otxtLlE.Value))
                {
                    DllFunciones.sendMessageBox(sboapp, "Por favor diligenciar la llave de la empresa");
                    return false;

                }
                else if (string.IsNullOrEmpty(otxtPwdE.Value))
                {
                    DllFunciones.sendMessageBox(sboapp, "Por favor diligenciar el Password de la Llave");
                    return false;

                }
                else if (string.IsNullOrEmpty(otxtMdo.Value))
                {
                    DllFunciones.sendMessageBox(sboapp, "Por favor diligenciar el modo de la base de datos");
                    return false;

                }
                else if (string.IsNullOrEmpty(otxtL.Value))
                {
                    DllFunciones.sendMessageBox(sboapp, "Por favor diligenciar la localizacion utilizada de la Llave");
                    return false;
                }
                else if (iCounterActiveConfig > 1)
                {
                    DllFunciones.sendMessageBox(sboapp, "Solo puede estar activar una parametrizacion a la vez, por favor inhabilite las demas parametrizaciones");
                    oChkStatus.Checked = false;
                    return false;
                }
                else
                {

                    if (oFormParametros.Mode == BoFormMode.fm_ADD_MODE)
                    {
                        sGetLastRecord = DllFunciones.GetStringXMLDocument(oCompany, "eBilling", "eBilling", "GetLastRecord");

                        oGetLastRecord.DoQuery(sGetLastRecord);
                        sGetLastRecord = oGetLastRecord.Fields.Item(0).Value.ToString();

                        oDocEntry.Value = sGetLastRecord;
                    }

                    DllFunciones.liberarObjetos(oGetLastRecord);
                    DllFunciones.liberarObjetos(oActiveConfig);
                    InsertDataSeriesNumber(oCompany, oFormParametros);

                    return true;
                }

                #endregion

            }
            catch (Exception)
            {
                throw;
            }

        }

        public Boolean Validate_oBusinessPartnerd(SAPbouiCOM.Application sboapp, SAPbobsCOM.Company oCompany, SAPbouiCOM.Form oFormBusinessPartnerd, string _sMotor)
        {
            Boolean Flag;

            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            try
            {
                #region Variables y objetos

                EditText BO_Email_1;
                EditText BO_RF;
                SAPbouiCOM.ComboBox BO_TR;

                BO_Email_1 = (EditText)oFormBusinessPartnerd.Items.Item("txtEmail1").Specific;
                BO_RF = (EditText)oFormBusinessPartnerd.Items.Item("txtRF").Specific;
                BO_TR = (SAPbouiCOM.ComboBox)oFormBusinessPartnerd.Items.Item("cboTR").Specific;

                #endregion

                #region Validación de los campos obligatorios 

                if (string.IsNullOrEmpty(BO_Email_1.Value))
                {
                    DllFunciones.sendMessageBox(sboapp, "Por favor diligenciar al menos 1 correo electronico en la pestaña 'eBilling'");
                    Flag = false;

                }
                else if (string.IsNullOrEmpty(BO_RF.Value))
                {
                    DllFunciones.sendMessageBox(sboapp, "Por favor seleccionar la responsabilidad fiscal en la pestaña 'eBilling'");
                    Flag = false;

                }
                else if (string.IsNullOrEmpty(BO_TR.Value))
                {
                    DllFunciones.sendMessageBox(sboapp, "Por favor diligenciar el Tipo de regimen en la pestaña 'eBilling'");
                    Flag = false;
                }
                else
                {
                    Flag = true;
                }

                #endregion

            }
            catch (Exception e)
            {

                DllFunciones.sendErrorMessage(sboapp, e);
                Flag = false;
            }

            return Flag;
        }

        public Boolean Validate_oInvoices(SAPbouiCOM.Application sboapp, SAPbobsCOM.Company oCompany, SAPbouiCOM.Form oForInvoices, string _sMotor)
        {
            Boolean Flag;

            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            try
            {
                #region Variables y objetos

                SAPbouiCOM.ComboBox BO_MP;

                BO_MP = (SAPbouiCOM.ComboBox)oForInvoices.Items.Item("txtMP").Specific;

                #endregion

                #region Validación de los campos obligatorios 

                if (string.IsNullOrEmpty(BO_MP.Value))
                {
                    DllFunciones.sendMessageBox(sboapp, "Por favor diligenciar el Medio de pago en la pestaña 'eBilling'");
                    Flag = false;

                }
                else
                {
                    Flag = true;
                }

                #endregion

            }
            catch (Exception e)
            {

                DllFunciones.sendErrorMessage(sboapp, e);
                Flag = false;
            }

            return Flag;
        }

        public Boolean Validate_oCreditNote(SAPbouiCOM.Application sboapp, SAPbobsCOM.Company oCompany, SAPbouiCOM.Form oForInvoices, string _sMotor)
        {
            Boolean Flag;

            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            try
            {
                #region Variables y objetos

                SAPbouiCOM.EditText BO_AFV;
                SAPbouiCOM.EditText BO_EBC;
                SAPbouiCOM.ComboBox BO_TN;

                BO_AFV = (SAPbouiCOM.EditText)oForInvoices.Items.Item("txtAFV").Specific;
                BO_EBC = (SAPbouiCOM.EditText)oForInvoices.Items.Item("txtEBC").Specific;
                BO_TN = (SAPbouiCOM.ComboBox)oForInvoices.Items.Item("txtTipN").Specific;

                #endregion

                #region Validación de los campos obligatorios 

                //if (string.IsNullOrEmpty(BO_AFV.Value))
                //{
                //    DllFunciones.sendMessageBox(sboapp, "Por favor diligenciar el Numero de la factura de venta en la pestaña 'eBilling'");
                //    Flag = false;

                //}
                if (string.IsNullOrEmpty(BO_EBC.Value))
                {
                    DllFunciones.sendMessageBox(sboapp, "Por favor diligenciar los comentarios 'Fac.Elec' en la pestaña 'eBilling'");
                    Flag = false;
                }
                if (string.IsNullOrEmpty(BO_TN.Value))
                {
                    DllFunciones.sendMessageBox(sboapp, "Por favor diligenciar el 'Tipo de Nota' en la pestaña 'eBilling'");
                    Flag = false;
                }
                else
                {
                    Flag = true;
                }

                #endregion

            }
            catch (Exception e)
            {

                DllFunciones.sendErrorMessage(sboapp, e);
                Flag = false;
            }

            return Flag;
        }

        public Boolean Validate_oDebitNote(SAPbouiCOM.Application sboapp, SAPbobsCOM.Company oCompany, SAPbouiCOM.Form oFormDebitNote, string _sMotor)
        {
            Boolean Flag;

            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            try
            {
                #region Variables y objetos

                SAPbouiCOM.EditText BO_AFV;
                SAPbouiCOM.EditText BO_EBC;
                SAPbouiCOM.ComboBox BO_TN;

                BO_AFV = (SAPbouiCOM.EditText)oFormDebitNote.Items.Item("txtAFV").Specific;
                BO_EBC = (SAPbouiCOM.EditText)oFormDebitNote.Items.Item("txtEBC").Specific;
                BO_TN = (SAPbouiCOM.ComboBox)oFormDebitNote.Items.Item("txtTipND").Specific;

                #endregion

                #region Validación de los campos obligatorios 

                if (string.IsNullOrEmpty(BO_AFV.Value))
                {
                    DllFunciones.sendMessageBox(sboapp, "Por favor diligenciar el Numero de la factura de venta en la pestaña 'eBilling'");
                    Flag = false;

                }
                if (string.IsNullOrEmpty(BO_EBC.Value))
                {
                    DllFunciones.sendMessageBox(sboapp, "Por favor diligenciar los comentarios 'Fac.Elec' en la pestaña 'eBilling'");
                    Flag = false;
                }
                if (string.IsNullOrEmpty(BO_TN.Value))
                {
                    DllFunciones.sendMessageBox(sboapp, "Por favor diligenciar el 'Tipo de Nota' en la pestaña 'eBilling'");
                    Flag = false;
                }
                else
                {
                    Flag = true;
                }

                #endregion

            }
            catch (Exception e)
            {

                DllFunciones.sendErrorMessage(sboapp, e);
                Flag = false;
            }

            return Flag;
        }

        public void ItemEvent_eBilling(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            switch (pVal.EventType)
            {
                case SAPbouiCOM.BoEventTypes.et_CLICK:
                    break;

                case BoEventTypes.et_FORM_ACTIVATE:
                    break;

                case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:

                    break;
            }

        }

        public void LinkedButtonMatrixFormVDBO(SAPbouiCOM.Application _sboapp, SAPbobsCOM.Company _oCompany, SAPbouiCOM.Form oFormVDBO, ItemEvent pVal, string _TipoDocumento)
        {
            if (_TipoDocumento == "FacturaDeClientes")
            {
                SAPbouiCOM.Matrix oMatrixInvoiceLinkedButton = (Matrix)oFormVDBO.Items.Item("MtxOINV").Specific;

                oFormVDBO.Freeze(true);

                SAPbouiCOM.Column oColumDocEntry;
                SAPbouiCOM.LinkedButton LkBtnDocEntry;

                oColumDocEntry = oMatrixInvoiceLinkedButton.Columns.Item("Col_9");

                LkBtnDocEntry = (SAPbouiCOM.LinkedButton)oColumDocEntry.ExtendedObject;
                LkBtnDocEntry.LinkedObject = BoLinkedObject.lf_Invoice;

                oFormVDBO.Freeze(false);

            }
            else if (_TipoDocumento == "NotaCreditoClientes")
            {
                SAPbouiCOM.Matrix oMatrixCreditNoteLinkedButton = (Matrix)oFormVDBO.Items.Item("MtxORIN").Specific;

                oFormVDBO.Freeze(true);

                SAPbouiCOM.Column oColumDocEntry;
                SAPbouiCOM.LinkedButton LkBtnDocEntry;

                oColumDocEntry = oMatrixCreditNoteLinkedButton.Columns.Item("Col_9");

                LkBtnDocEntry = (SAPbouiCOM.LinkedButton)oColumDocEntry.ExtendedObject;
                LkBtnDocEntry.LinkedObject = BoLinkedObject.lf_InvoiceCreditMemo;

                oFormVDBO.Freeze(false);

            }
            else
            {
                SAPbouiCOM.Matrix oMatrixInvoiceLinkedButton = (Matrix)oFormVDBO.Items.Item("MtxOINVD").Specific;

                oFormVDBO.Freeze(true);

                SAPbouiCOM.Column oColumDocEntry;
                SAPbouiCOM.LinkedButton LkBtnDocEntry;

                oColumDocEntry = oMatrixInvoiceLinkedButton.Columns.Item("Col_9");

                LkBtnDocEntry = (SAPbouiCOM.LinkedButton)oColumDocEntry.ExtendedObject;
                LkBtnDocEntry.LinkedObject = BoLinkedObject.lf_Invoice;

                oFormVDBO.Freeze(false);

            }

        }

        public void EnviarDocumentoTFHKA(SAPbouiCOM.Application _sboapp, SAPbobsCOM.Company _oCompany, SAPbouiCOM.Form _oFormInvoices, string _TipoDocumento, string TipoIntegracion)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            if (TipoIntegracion == "A")
            {
                if (_oFormInvoices.Mode == BoFormMode.fm_OK_MODE)
                {
                    #region Envio del documento por el Modulo 

                    #region Consulta URL

                    string sGetModo = null;
                    string sURLEmision = null;
                    string sURLAdjuntos = null;
                    string sModo = null;

                    SAPbobsCOM.Recordset oConsultarGetModo = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    sGetModo = DllFunciones.GetStringXMLDocument(_oCompany, "eBilling", "eBilling", "GetModoandURL");

                    sGetModo = sGetModo.Replace("%Estado%", "\"U_BO_Status\" = 'Y'").Replace("%DocEntry%", " ");

                    oConsultarGetModo.DoQuery(sGetModo);

                    sURLEmision = Convert.ToString(oConsultarGetModo.Fields.Item("URLTFHKA").Value.ToString()) + "/ws/v1.0/Service.svc?wsdl";
                    sURLAdjuntos = Convert.ToString(oConsultarGetModo.Fields.Item("URLTFHKA").Value.ToString()) + "/ws/adjuntos/Service.svc?wsdl";
                    sModo = Convert.ToString(oConsultarGetModo.Fields.Item("Modo").Value.ToString());

                    DllFunciones.liberarObjetos(oConsultarGetModo);

                    #endregion

                    #region Instanciacion parametros TFHKA

                    //Especifica el puerto (HTTP o HTTPS)
                    if (sModo == "PRU")
                    {
                        BasicHttpBinding port = new BasicHttpBinding();
                    }
                    else if (sModo == "PRO")
                    {
                        BasicHttpsBinding port = new BasicHttpsBinding();
                    }

                    port.MaxBufferPoolSize = Int32.MaxValue;
                    port.MaxBufferSize = Int32.MaxValue;
                    port.MaxReceivedMessageSize = Int32.MaxValue;
                    port.ReaderQuotas.MaxStringContentLength = Int32.MaxValue;
                    port.SendTimeout = TimeSpan.FromMinutes(2);
                    port.ReceiveTimeout = TimeSpan.FromMinutes(2);

                    if (sModo == "PRO")
                    {
                        port.Security.Mode = BasicHttpSecurityMode.Transport;
                    }

                    //Especifica la dirección de conexion para Emision y Adjuntos 
                    EndpointAddress endPointEmision = new EndpointAddress(sURLEmision); //URL DEMO EMISION
                    EndpointAddress endPointAdjuntos = new EndpointAddress(sURLAdjuntos); //URL DEMO ADJUNTOS          

                    #endregion

                    #region Variables

                    string sDocNumInvoice = null;
                    string sSerieNumeracion = null;
                    string sQueryDocEntryDocument = null;
                    string sDocEntryInvoice = null;
                    string sProcedureXML = null;
                    string sDocumentoCabecera = null;
                    string sDocumentoLinea = null;
                    string sDocumentoImpuestosGenerales = null;
                    string sDocumentoImpuestosTotales = null;
                    string sParametrosTFHKA = null;
                    string sRutaCR = null;
                    string sPrefijoConDoc = null;
                    string sPrefijo = null;
                    string sStatusDoc = null;
                    string sFormaEnvio = null;
                    string sLlave = null;
                    string sPassword = null;
                    string sUserDB = null;
                    string sPassDB = null;
                    string sRutaPDF = null;
                    string sRutaXML = null;
                    string sNombreDocumento = null;
                    string sNombreDocWarning = null;
                    string sCUFEInvoice = null;
                    int sReprocesar = 0;
                    string sGenerarXMLPrueba = null;
                    string sCountsEmails = null;

                    Boolean GeneroPDF = false;

                    if (_TipoDocumento == "FacturaDeClientes")
                    {
                        sNombreDocumento = "Factura_de_Venta_No_";
                        sNombreDocWarning = "Factura de venta";
                    }
                    else if (_TipoDocumento == "NotaCreditoClientes")
                    {
                        sNombreDocumento = "Nota_Credito_No_";
                        sNombreDocWarning = "Nota credito de clientes";
                    }
                    else if (_TipoDocumento == "NotaDebitoClientes")
                    {
                        sNombreDocumento = "Nota_debito_Clientes_No_";
                        sNombreDocWarning = "Nota debito de clientes";
                    }

                    #endregion

                    DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 1: Consultando " + sNombreDocWarning + " ...");

                    #region Consulta de documento en la base de datos y el estado del documento

                    sDocNumInvoice = ((SAPbouiCOM.EditText)(_oFormInvoices.Items.Item("8").Specific)).Value.ToString();
                    SAPbouiCOM.ComboBox cbSerieNumeracion = (SAPbouiCOM.ComboBox)(_oFormInvoices.Items.Item("88").Specific);
                    sSerieNumeracion = cbSerieNumeracion.Selected.Value;

                    SAPbobsCOM.Recordset oConsultaDocEntry = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    sQueryDocEntryDocument = DllFunciones.GetStringXMLDocument(_oCompany, "eBilling", "eBilling", "GetDocEntryAndParameters");

                    if (_TipoDocumento == "FacturaDeClientes")
                    {
                        sQueryDocEntryDocument = sQueryDocEntryDocument.Replace("%sDocNumInvoice%", sDocNumInvoice).Replace("%sSerieNumeracion%", sSerieNumeracion).Replace("%Tabla%", "OINV").Replace("%DocSubType%", "--");
                    }
                    else if (_TipoDocumento == "NotaCreditoClientes")
                    {
                        sQueryDocEntryDocument = sQueryDocEntryDocument.Replace("%sDocNumInvoice%", sDocNumInvoice).Replace("%sSerieNumeracion%", sSerieNumeracion).Replace("%Tabla%", "ORIN").Replace("%DocSubType%", "--");
                    }
                    else if (_TipoDocumento == "NotaDebitoClientes")
                    {
                        sQueryDocEntryDocument = sQueryDocEntryDocument.Replace("%sDocNumInvoice%", sDocNumInvoice).Replace("%sSerieNumeracion%", sSerieNumeracion).Replace("%Tabla%", "OINV").Replace("%DocSubType%", "DN");
                    }

                    oConsultaDocEntry.DoQuery(sQueryDocEntryDocument);

                    sPrefijo = Convert.ToString(oConsultaDocEntry.Fields.Item("PrefijoDes").Value.ToString());
                    sStatusDoc = Convert.ToString(oConsultaDocEntry.Fields.Item("CRWS").Value.ToString());
                    sFormaEnvio = Convert.ToString(oConsultaDocEntry.Fields.Item("FormaEnvio").Value.ToString());
                    sLlave = Convert.ToString(oConsultaDocEntry.Fields.Item("Llave").Value.ToString());
                    sPassword = Convert.ToString(oConsultaDocEntry.Fields.Item("Password").Value.ToString());
                    sUserDB = Convert.ToString(oConsultaDocEntry.Fields.Item("UserDB").Value.ToString());
                    sPassDB = Convert.ToString(oConsultaDocEntry.Fields.Item("PassDB").Value.ToString());
                    sRutaXML = Convert.ToString(oConsultaDocEntry.Fields.Item("RutaXML").Value.ToString()) + "\\" + sNombreDocumento + sPrefijo + '_' + sDocNumInvoice + ".txt";
                    sRutaPDF = Convert.ToString(oConsultaDocEntry.Fields.Item("RutaPDF").Value.ToString()) + "\\" + sNombreDocumento + sPrefijo + '_' + sDocNumInvoice + ".pdf";
                    sRutaCR = Convert.ToString(oConsultaDocEntry.Fields.Item("RutaCR").Value.ToString());
                    sGenerarXMLPrueba = Convert.ToString(oConsultaDocEntry.Fields.Item("GeneraXMLP").Value.ToString());
                    sCountsEmails = DllFunciones.GetStringXMLDocument(_oCompany, "eBilling", "eBilling", "CountsEmails");

                    #endregion

                    if (sStatusDoc == "200")
                    {
                        #region Si el estado del documento es 200 Pregunta al usuario si desea volver aenviar la factura a la DIAN, 

                        sReprocesar = DllFunciones.sendMessageBoxY_N(_sboapp, "La " + sNombreDocWarning + " ya fue emitida a la DIAN, ¿ Desea volver a enviarla ?");

                        if (sReprocesar == 1)
                        {
                            if (oConsultaDocEntry.RecordCount > 0)
                            {

                                DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 2: Creando Objeto " + sNombreDocWarning + " ...");

                                #region Si existe el numero de Documento, busca y crea el objeto factura

                                sDocEntryInvoice = oConsultaDocEntry.Fields.Item(0).Value.ToString();

                                SAPbobsCOM.Recordset oCabeceraDocumento = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                SAPbobsCOM.Recordset oLineasDocumento = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                SAPbobsCOM.Recordset oImpuestosGenerales = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                SAPbobsCOM.Recordset oImpuestosTotales = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                SAPbobsCOM.Recordset oCUFEInvoice = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                                sProcedureXML = DllFunciones.GetStringXMLDocument(_oCompany, "eBilling", "eBilling", "ExecProcedureBOFacturaXML");

                                if (_TipoDocumento == "FacturaDeClientes")
                                {
                                    sDocumentoCabecera = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "13").Replace("%TipoConsulta%", "Encabezado");
                                    sDocumentoLinea = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "13").Replace("%TipoConsulta%", "Lineas");
                                    sDocumentoImpuestosGenerales = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "13").Replace("%TipoConsulta%", "Impuestos");
                                    sDocumentoImpuestosTotales = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "13").Replace("%TipoConsulta%", "ImpuestosTotales");
                                }
                                else if (_TipoDocumento == "NotaCreditoClientes")
                                {
                                    sDocumentoCabecera = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "14").Replace("%TipoConsulta%", "Encabezado");
                                    sDocumentoLinea = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "14").Replace("%TipoConsulta%", "Lineas");
                                    sDocumentoImpuestosGenerales = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "14").Replace("%TipoConsulta%", "Impuestos");
                                    sDocumentoImpuestosTotales = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "14").Replace("%TipoConsulta%", "ImpuestosTotales");
                                }
                                else if (_TipoDocumento == "NotaDebitoClientes")
                                {
                                    sDocumentoCabecera = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "13_ND").Replace("%TipoConsulta%", "Encabezado");
                                    sDocumentoLinea = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "13_ND").Replace("%TipoConsulta%", "Lineas");
                                    sDocumentoImpuestosGenerales = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "13_ND").Replace("%TipoConsulta%", "Impuestos");
                                    sDocumentoImpuestosTotales = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "13_ND").Replace("%TipoConsulta%", "ImpuestosTotales");
                                }

                                oCabeceraDocumento.DoQuery(sDocumentoCabecera);
                                oLineasDocumento.DoQuery(sDocumentoLinea);
                                oImpuestosGenerales.DoQuery(sDocumentoImpuestosGenerales);
                                oImpuestosTotales.DoQuery(sDocumentoImpuestosTotales);

                                if (_TipoDocumento == "NotaCreditoClientes")
                                {
                                    sCUFEInvoice = DllFunciones.GetStringXMLDocument(_oCompany, "eBilling", "eBilling", "GetCUFEInvoice");
                                    sCUFEInvoice = sCUFEInvoice.Replace("%DocNum%", Convert.ToString(oCabeceraDocumento.Fields.Item("No_FV").Value.ToString()));

                                    oCUFEInvoice.DoQuery(sCUFEInvoice);
                                }
                                else if (_TipoDocumento == "NotaDebitoClientes")
                                {
                                    sCUFEInvoice = DllFunciones.GetStringXMLDocument(_oCompany, "eBilling", "eBilling", "GetCUFEDebitNote");
                                    sCUFEInvoice = sCUFEInvoice.Replace("%DocNum%", Convert.ToString(oCabeceraDocumento.Fields.Item("No_FV").Value.ToString()));

                                    oCUFEInvoice.DoQuery(sCUFEInvoice);

                                }
                                FacturaGeneral Documento = oBuillInvoice(oCabeceraDocumento, oLineasDocumento, oImpuestosGenerales, oImpuestosTotales, oCUFEInvoice, _TipoDocumento);

                                #endregion

                                DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 3: Guardando TXT " + sNombreDocWarning + " ...");

                                #region Guarda el TXT en la ruta del XML configurada

                                StreamWriter MyFile = new StreamWriter(sRutaXML); //ruta y name del archivo request a almecenar


                                #endregion

                                DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 4: Serealizando la " + sNombreDocWarning + " ...");

                                #region Serealizando el documento

                                SAPbobsCOM.Recordset oParametrosTFHKA = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                                sParametrosTFHKA = DllFunciones.GetStringXMLDocument(_oCompany, "eBilling", "eBilling", "GetParameterstoSend");

                                oParametrosTFHKA.DoQuery(sParametrosTFHKA);

                                XmlSerializer Serializer1 = new XmlSerializer(typeof(FacturaGeneral));
                                Serializer1.Serialize(MyFile, Documento); // Objeto serializado
                                MyFile.Close();

                                if (sGenerarXMLPrueba == "N")
                                {
                                    File.Delete(sRutaXML);
                                }

                                #endregion

                                DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 5: Enviando " + sNombreDocWarning + " a TFHKA...");

                                #region Envio del objeto factura a TFHKA

                                serviceClient = new eBilling.ServicioEmisionFE.ServiceClient(port, endPointEmision);
                                serviceClientAdjuntos = new eBilling.ServicioAdjuntosFE.ServiceClient(port, endPointAdjuntos);

                                DocumentResponse RespuestaDoc = new eBilling.ServicioEmisionFE.DocumentResponse(); //objeto Response del metodo enviar

                                if (string.IsNullOrEmpty(sLlave))
                                {
                                    DllFunciones.sendMessageBox(_sboapp, "Error Paso 4: No se ha parametrizado la llave de TFHKA en la configuracion Inicial, por lo cual no se puede enviar la factura a la DIAN ");
                                }
                                else if (string.IsNullOrEmpty(sPassword))
                                {
                                    DllFunciones.sendMessageBox(_sboapp, "Error Paso 4: No se ha parametrizado el password de TFHKA en la configuracion Inicial, por lo cual no se puede enviar la factura a la DIAN");
                                }
                                else
                                {
                                    #region Respuesta el Web Service de TFHKA y actualizacion de los campos en la factura

                                    RespuestaDoc = serviceClient.Enviar(sLlave, sPassword, Documento, sFormaEnvio);

                                    if (RespuestaDoc.codigo == 200)
                                    {
                                        #region Procesa la repuesta

                                        DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 6: " + sNombreDocWarning + " enviada correctamente a TFHKA");

                                        #region Se actualiza el documento en SAP con las respuesta de TFHKA

                                        if (_TipoDocumento == "FacturaDeClientes" || (_TipoDocumento == "NotaDebitoClientes"))
                                        {
                                            UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, "Documento autorizado por la DIAN", RespuestaDoc.cufe, RespuestaDoc.qr, null, null);
                                            InsertSendEmail(_oCompany, oCabeceraDocumento, sCountsEmails, sDocEntryInvoice, "13");

                                        }
                                        else if (_TipoDocumento == "NotaCreditoClientes")
                                        {
                                            UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, "Documento autorizado por la DIAN", RespuestaDoc.cufe, RespuestaDoc.qr, null, null);
                                            InsertSendEmail(_oCompany, oCabeceraDocumento, sCountsEmails, sDocEntryInvoice, "14");

                                        }

                                        #endregion

                                        #region Valida la forma de envio,si es 11,  genera el PDF y retorna confirmacion de la generacion del PDF

                                        if (sFormaEnvio == "11")
                                        {
                                            DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 7: Generando PDF, por favor espere ...");

                                            FileInfo ValidacionPDF = new FileInfo(sRutaPDF);

                                            if (ValidacionPDF.Exists)
                                            {
                                                GeneroPDF = true;
                                            }
                                            else
                                            {
                                                GeneroPDF = ExportPDF(_sboapp, _oCompany, sRutaPDF, sDocEntryInvoice, sRutaCR, _TipoDocumento, sUserDB, sPassDB);
                                            }
                                        }

                                        #endregion

                                        #region Si genera correctamente el PDF lo adjunta a la factura de venta en SAP, 

                                        if (GeneroPDF == true)
                                        {
                                            DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 7: Adjuntando PDF, por favor espere ...");

                                            if (_TipoDocumento == "FacturaDeClientes" || _TipoDocumento == "NotaDebitoClientes")
                                            {
                                                UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, 0, null, null, null, sRutaPDF, null);
                                            }
                                            else if (_TipoDocumento == "NotaCreditoClientes")
                                            {
                                                UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, 0, null, null, null, sRutaPDF, null);
                                            }

                                        }
                                        else
                                        {
                                        }

                                        #endregion

                                        DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 8: Enviando PDF a TFHKA por favor espere ...");

                                        #region Envia el PDF al proveedor tecnologico TFHKA

                                        sPrefijoConDoc = Convert.ToString(oCabeceraDocumento.Fields.Item("Prefijo").Value.ToString()) + sDocNumInvoice;

                                        EnviarAdjuntosTFHKA(_sboapp, oCabeceraDocumento, sRutaPDF, sPrefijoConDoc, sLlave, sPassword);

                                        #endregion

                                        #region Se descarga el XML y se adjunta a la factura de venta

                                        DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 7: Adjuntando XML y actualizando campos, por favor espere ...");

                                        #region Descarga el XML y retorna la confirmacion

                                        bool DescargoXML = false;

                                        DescargoXML = DescargaXML(_oCompany, sPrefijoConDoc, Convert.ToString(oParametrosTFHKA.Fields.Item("TokenEmpresa").Value.ToString()), Convert.ToString(oParametrosTFHKA.Fields.Item("TokenPassword").Value.ToString()), sRutaXML);

                                        #endregion

                                        #region Actualiza el campo de XML en el documento de SAP

                                        if (DescargoXML == true)
                                        {

                                            if (_TipoDocumento == "FacturaDeClientes" || _TipoDocumento == "NotaDebitoClientes")
                                            {
                                                UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, 0, null, null, null, null, sRutaXML.Replace(".txt", ".xml"));

                                            }
                                            else if (_TipoDocumento == "NotaCreditoClientes")
                                            {
                                                UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, 0, null, null, null, sRutaPDF, null);
                                            }

                                        }
                                        else
                                        {

                                        }

                                        #endregion

                                        DllFunciones.sendMessageBox(_sboapp, "Codigo de mensaje No. " + RespuestaDoc.codigo.ToString() + ", " + "Documento enviado correctamente");

                                        _sboapp.ActivateMenuItem("1304");

                                        #endregion

                                        DllFunciones.sendMessageBox(_sboapp, "El documento fue enviado existosamente a la DIAN");

                                        DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 9: Proceso de emision Finalizado ...");

                                        _sboapp.ActivateMenuItem("1304");

                                        #endregion
                                    }
                                    else if (RespuestaDoc.codigo == 201)
                                    {
                                        #region Procesa la respuesta                             

                                        DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 6: " + sNombreDocWarning + " enviada correctamente a TFHKA");

                                        #region Consulta el estado del documento en TFHKA

                                        sPrefijoConDoc = Convert.ToString(oCabeceraDocumento.Fields.Item("Prefijo").Value.ToString()) + sDocNumInvoice;
                                        DocumentStatusResponse resp = serviceClient.EstadoDocumento(Convert.ToString(oParametrosTFHKA.Fields.Item("TokenEmpresa").Value.ToString()), Convert.ToString(oParametrosTFHKA.Fields.Item("TokenPassword").Value.ToString()), sPrefijoConDoc);

                                        #endregion

                                        #region Actualiza el documento con la respuesta de TFHKA

                                        if (_TipoDocumento == "FacturaDeClientes" || (_TipoDocumento == "NotaDebitoClientes"))
                                        {
                                            UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, resp.codigo, "Documento autorizado por la DIAN", resp.cufe, resp.cadenaCodigoQR, null, null);
                                            InsertSendEmail(_oCompany, oCabeceraDocumento, sCountsEmails, sDocEntryInvoice, "13");
                                        }
                                        else if (_TipoDocumento == "NotaCreditoClientes")
                                        {
                                            UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, resp.codigo, "Documento autorizado por la DIAN", resp.cufe, resp.cadenaCodigoQR, null, null);
                                            InsertSendEmail(_oCompany, oCabeceraDocumento, sCountsEmails, sDocEntryInvoice, "14");
                                        }
                                        #endregion

                                        if (resp.codigo == 200)
                                        {
                                            #region Valida la forma de envio,si es 11,  genera el PDF y retorna confirmacion de la generacion del PDF

                                            if (sFormaEnvio == "11")
                                            {
                                                DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 7: Generando PDF, por favor espere ...");

                                                FileInfo ValidacionPDF = new FileInfo(sRutaPDF);

                                                if (ValidacionPDF.Exists)
                                                {
                                                    GeneroPDF = true;
                                                }
                                                else
                                                {
                                                    GeneroPDF = ExportPDF(_sboapp, _oCompany, sRutaPDF, sDocEntryInvoice, sRutaCR, _TipoDocumento, sUserDB, sPassDB);
                                                }

                                            }
                                            else
                                            {
                                                DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 8: Proceso de emision finalizado con exito.");
                                            }

                                            DllFunciones.sendMessageBox(_sboapp, "El documento fue enviado existosamente a la DIAN");

                                            #endregion
                                        }

                                        _sboapp.ActivateMenuItem("1304");

                                        #endregion
                                    }
                                    else if (RespuestaDoc.codigo == 101)
                                    {
                                        #region Procesa la respuesta

                                        if (_TipoDocumento == "FacturaDeClientes" || (_TipoDocumento == "NotaDebitoClientes"))
                                        {
                                            UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, RespuestaDoc.mensaje, "", "", null, null);
                                        }
                                        else if (_TipoDocumento == "NotaCreditoClientes")
                                        {
                                            UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, RespuestaDoc.mensaje, "", "", null, null);
                                        }

                                        DllFunciones.sendMessageBox(_sboapp, "Codigo de error No. " + RespuestaDoc.codigo.ToString() + ", " + RespuestaDoc.mensaje.ToString());

                                        _sboapp.ActivateMenuItem("1304");

                                        #endregion
                                    }
                                    else if (RespuestaDoc.codigo == 99)
                                    {
                                        #region Procesa la respuesta

                                        if (_TipoDocumento == "FacturaDeClientes" || (_TipoDocumento == "NotaDebitoClientes"))
                                        {
                                            UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, Convert.ToString(RespuestaDoc.reglasValidacionDIAN.GetValue(0)), "", "", null, null);
                                        }
                                        else if (_TipoDocumento == "NotaCreditoClientes")
                                        {
                                            UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, Convert.ToString(RespuestaDoc.reglasValidacionDIAN.GetValue(0)), "", "", null, null);
                                        }

                                        DllFunciones.sendMessageBox(_sboapp, "Codigo de error No. " + RespuestaDoc.codigo.ToString() + ", " + RespuestaDoc.reglasValidacionDIAN.ToString());

                                        _sboapp.ActivateMenuItem("1304");

                                        #endregion
                                    }

                                    else if (RespuestaDoc.codigo == 109)
                                    {
                                        #region Procesa la respuesta

                                        if (_TipoDocumento == "FacturaDeClientes" || (_TipoDocumento == "NotaDebitoClientes"))
                                        {
                                            UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, Convert.ToString(RespuestaDoc.mensajesValidacion.GetValue(0)), "", "", null, null);
                                        }
                                        else if (_TipoDocumento == "NotaCreditoClientes")
                                        {
                                            UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, Convert.ToString(RespuestaDoc.mensajesValidacion.GetValue(0)), "", "", null, null);
                                        }

                                        DllFunciones.sendMessageBox(_sboapp, "Codigo de error No. " + RespuestaDoc.codigo.ToString() + ", " + RespuestaDoc.mensaje.ToString() + " " + RespuestaDoc.mensajesValidacion.GetValue(0));

                                        _sboapp.ActivateMenuItem("1304");

                                        #endregion
                                    }
                                    else if (RespuestaDoc.codigo == 110)
                                    {
                                        #region Procesa la respuesta

                                        if (_TipoDocumento == "FacturaDeClientes" || (_TipoDocumento == "NotaDebitoClientes")) if (_TipoDocumento == "FacturaDeClientes")
                                            {
                                                UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, RespuestaDoc.mensaje.ToString() + ", Total de Factura es diferente de la suma de Total valor bruto + Tributos - Total Tributo Retenidos - Anticipos ", "", "", null, null);
                                            }
                                            else if (_TipoDocumento == "NotaCreditoClientes")
                                            {
                                                UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, RespuestaDoc.mensaje.ToString() + ", Total de Factura es diferente de la suma de Total valor bruto + Tributos - Total Tributo Retenidos - Anticipos ", "", "", null, null);
                                            }

                                        DllFunciones.sendMessageBox(_sboapp, "Codigo de error No. " + RespuestaDoc.codigo.ToString() + ", " + RespuestaDoc.mensaje.ToString() + ", Total de Factura es diferente de la suma de Total valor bruto + Tributos - Total Tributo Retenidos - Anticipos ");

                                        _sboapp.ActivateMenuItem("1304");

                                        #endregion
                                    }
                                    else if (RespuestaDoc.codigo == 111)
                                    {
                                        #region Procesa la respuesta

                                        if (_TipoDocumento == "FacturaDeClientes" || (_TipoDocumento == "NotaDebitoClientes"))
                                        {
                                            UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, RespuestaDoc.mensaje.ToString(), "", "", null, null);
                                        }
                                        else if (_TipoDocumento == "NotaCreditoClientes")
                                        {
                                            UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, RespuestaDoc.mensaje.ToString(), "", "", null, null);
                                        }

                                        DllFunciones.sendMessageBox(_sboapp, "Codigo de error No. " + RespuestaDoc.codigo.ToString() + ", " + RespuestaDoc.mensaje.ToString());

                                        _sboapp.ActivateMenuItem("1304");

                                        #endregion

                                    }
                                    else if (RespuestaDoc.codigo == 112)
                                    {
                                        #region Procesa la respuesta

                                        if (_TipoDocumento == "FacturaDeClientes" || (_TipoDocumento == "NotaDebitoClientes"))
                                        {
                                            UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, RespuestaDoc.mensaje.ToString(), "", "", null, null);
                                        }
                                        else if (_TipoDocumento == "NotaCreditoClientes")
                                        {
                                            UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, RespuestaDoc.mensaje.ToString(), "", "", null, null);
                                        }

                                        DllFunciones.sendMessageBox(_sboapp, "Codigo de error No. " + RespuestaDoc.codigo.ToString() + ", " + RespuestaDoc.mensaje.ToString());

                                        _sboapp.ActivateMenuItem("1304");

                                        #endregion
                                    }
                                    else if (RespuestaDoc.codigo == 150)
                                    {
                                        #region Procesa la respuesta

                                        DllFunciones.sendMessageBox(_sboapp, "Codigo de error No. " + RespuestaDoc.codigo.ToString() + ", " + RespuestaDoc.mensaje.ToString());

                                        #endregion
                                    }
                                    else if (RespuestaDoc.codigo == 114)
                                    {
                                        #region Procesa la respuesta

                                        DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 5: Consultando el estado del documento a TFHKA, por favor espere ...");

                                        #region Consulta el estado del documento en el proveedor tecnologico

                                        sPrefijoConDoc = Convert.ToString(oCabeceraDocumento.Fields.Item("Prefijo").Value.ToString()) + sDocNumInvoice;
                                        DocumentStatusResponse resp = new eBilling.ServicioEmisionFE.DocumentStatusResponse();
                                        resp = serviceClient.EstadoDocumento(sLlave, sPassword, sPrefijoConDoc);

                                        #endregion

                                        #region Se actualiza la factura con las respuesta de TFHKA

                                        if (_TipoDocumento == "FacturaDeClientes" || (_TipoDocumento == "NotaDebitoClientes"))
                                        {
                                            UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, resp.codigo, "Documento autorizado por la DIAN", resp.cufe, resp.cadenaCodigoQR, null, null);
                                            InsertSendEmail(_oCompany, oCabeceraDocumento, sCountsEmails, sDocEntryInvoice, "13");
                                        }
                                        else if (_TipoDocumento == "NotaCreditoClientes")
                                        {
                                            UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, resp.codigo, "Documento autorizado por la DIAN", resp.cufe, resp.cadenaCodigoQR, null, null);
                                            InsertSendEmail(_oCompany, oCabeceraDocumento, sCountsEmails, sDocEntryInvoice, "14");
                                        }

                                        #endregion

                                        #region Valida la forma de envio, si es 11 genera el PDF y retorna confirmacion de la generacion del PDF

                                        if (sFormaEnvio == "11")
                                        {
                                            DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 6: Generando PDF, por favor espere ...");

                                            FileInfo ValidacionPDF = new FileInfo(sRutaPDF);

                                            if (ValidacionPDF.Exists)
                                            {
                                                GeneroPDF = true;
                                            }
                                            else
                                            {
                                                GeneroPDF = ExportPDF(_sboapp, _oCompany, sRutaPDF, sDocEntryInvoice, sRutaCR, _TipoDocumento, sUserDB, sPassDB);
                                            }

                                        }

                                        #endregion

                                        #region Si genera correctamente el PDF lo adjunta a la factura de venta en SAP, 

                                        if (GeneroPDF == true)
                                        {
                                            DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 7: Adjuntando PDF y actualizando campos, por favor espere ...");

                                            if (_TipoDocumento == "FacturaDeClientes" || _TipoDocumento == "NotaDebitoClientes")
                                            {
                                                UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, resp.codigo, null, resp.cufe, resp.cadenaCodigoQR, sRutaPDF, null);
                                            }
                                            else if (_TipoDocumento == "NotaCreditoClientes")
                                            {
                                                UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, resp.codigo, null, resp.cufe, resp.cadenaCodigoQR, sRutaPDF, null);
                                            }

                                            #region Envia el PDF al proveedor tecnologico TFHKA

                                            DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 8: Enviando PDF a TFHKA por favor espere ...");

                                            EnviarAdjuntosTFHKA(_sboapp, oCabeceraDocumento, sRutaPDF, sPrefijoConDoc, sLlave, sPassword);

                                            #endregion

                                            DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 9: Proceso de emision Finalizado ...");

                                        }
                                        else
                                        {
                                            DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 7: Actualizando campos, por favor espere ...");

                                            UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, 0, null, null, null, null, null);

                                            DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 8: Proceso de emision Finalizado ...");
                                        }

                                        #endregion

                                        #region Se descarga el XML y se adjunta a la factura de venta

                                        DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 7: Adjuntando XML y actualizando campos, por favor espere ...");

                                        #region Descarga el XML y retorna la confirmacion

                                        bool DescargoXML = false;

                                        DescargoXML = DescargaXML(_oCompany, sPrefijoConDoc, sLlave, sPassword, sRutaXML);

                                        #endregion

                                        #region Actualiza el campo de XML en el documento de SAP

                                        if (DescargoXML == true)
                                        {

                                            if (_TipoDocumento == "FacturaDeClientes" || _TipoDocumento == "NotaDebitoClientes")
                                            {
                                                UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, 0, null, null, null, null, sRutaXML.Replace(".txt", ".xml"));

                                            }
                                            else if (_TipoDocumento == "NotaCreditoClientes")
                                            {
                                                UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, 0, null, null, null, null, sRutaXML.Replace(".txt", ".xml"));
                                            }

                                        }
                                        else
                                        {

                                        }

                                        #endregion

                                        DllFunciones.sendMessageBox(_sboapp, "Codigo de mensaje No. " + RespuestaDoc.codigo.ToString() + ", " + "Documento enviado correctamente");

                                        _sboapp.ActivateMenuItem("1304");

                                        #endregion

                                        #endregion
                                    }
                                    else if (RespuestaDoc.codigo == 1)
                                    {
                                        #region Procesa la respuesta

                                        DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 5: Consultando el estado del documento a TFHKA, por favor espere ...");

                                        #region Consulta el estado del documento en el proveedor tecnologico

                                        sPrefijoConDoc = Convert.ToString(oCabeceraDocumento.Fields.Item("Prefijo").Value.ToString()) + sDocNumInvoice;
                                        DocumentStatusResponse resp = serviceClient.EstadoDocumento(Convert.ToString(oParametrosTFHKA.Fields.Item("TokenEmpresa").Value.ToString()), Convert.ToString(oParametrosTFHKA.Fields.Item("TokenPassword").Value.ToString()), sPrefijoConDoc);

                                        #endregion

                                        #region Se actualiza la factura con las respuesta de TFHKA

                                        if (_TipoDocumento == "FacturaDeClientes" || (_TipoDocumento == "NotaDebitoClientes"))
                                        {
                                            UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, resp.codigo, "Documento autorizado por la DIAN", resp.cufe, resp.cadenaCodigoQR, null, null);
                                            InsertSendEmail(_oCompany, oCabeceraDocumento, sCountsEmails, sDocEntryInvoice, "13");
                                        }
                                        else if (_TipoDocumento == "NotaCreditoClientes")
                                        {
                                            UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, resp.codigo, "Documento autorizado por la DIAN", resp.cufe, resp.cadenaCodigoQR, null, null);
                                            InsertSendEmail(_oCompany, oCabeceraDocumento, sCountsEmails, sDocEntryInvoice, "14");
                                        }

                                        #endregion

                                        #region Valida la forma de envio, si es 11 genera el PDF y retorna confirmacion de la generacion del PDF

                                        if (sFormaEnvio == "11")
                                        {
                                            DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 6: Generando PDF, por favor espere ...");

                                            FileInfo ValidacionPDF = new FileInfo(sRutaPDF);

                                            if (ValidacionPDF.Exists)
                                            {
                                                GeneroPDF = true;
                                            }
                                            else
                                            {
                                                GeneroPDF = ExportPDF(_sboapp, _oCompany, sRutaPDF, sDocEntryInvoice, sRutaCR, _TipoDocumento, sUserDB, sPassDB);
                                            }

                                        }

                                        #endregion

                                        #region Si genera correctamente el PDF lo adjunta a la factura de venta en SAP, 

                                        if (GeneroPDF == true)
                                        {
                                            DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 7: Adjuntando PDF y actualizando campos, por favor espere ...");

                                            if (_TipoDocumento == "FacturaDeClientes")
                                            {
                                                UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, resp.codigo, "Documento autorizado por la DIAN", resp.cufe, resp.cadenaCodigoQR, sRutaPDF, null);
                                            }
                                            else if (_TipoDocumento == "NotaCreditoClientes")
                                            {
                                                UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, resp.codigo, "Documento autorizado por la DIAN", resp.cufe, resp.cadenaCodigoQR, sRutaPDF, null);
                                            }

                                            #region Envia el PDF al proveedor tecnologico TFHKA

                                            DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 8: Enviando PDF a TFHKA por favor espere ...");

                                            EnviarAdjuntosTFHKA(_sboapp, oCabeceraDocumento, sRutaPDF, sPrefijoConDoc, sLlave, sPassword);

                                            #endregion

                                            DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 9: Proceso de emision Finalizado ...");

                                        }
                                        else
                                        {
                                            DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 7: Actualizando campos, por favor espere ...");

                                            UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, 0, null, null, null, null, null);

                                            DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 8: Proceso de emision Finalizado ...");
                                        }

                                        #endregion

                                        #region Se descarga el XML y se adjunta a la factura de venta

                                        DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 7: Adjuntando XML y actualizando campos, por favor espere ...");

                                        #region Descarga el XML y retorna la confirmacion

                                        bool DescargoXML = false;

                                        DescargoXML = DescargaXML(_oCompany, sPrefijoConDoc, Convert.ToString(oParametrosTFHKA.Fields.Item("TokenEmpresa").Value.ToString()), Convert.ToString(oParametrosTFHKA.Fields.Item("TokenPassword").Value.ToString()), sRutaXML);

                                        #endregion

                                        #region Actualiza el campo de XML en el documento de SAP

                                        if (DescargoXML == true)
                                        {

                                            if (_TipoDocumento == "FacturaDeClientes")
                                            {
                                                UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, 0, null, null, null, null, sRutaXML.Replace(".txt", ".xml"));

                                            }
                                            else if (_TipoDocumento == "NotaCreditoClientes")
                                            {
                                                UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, 0, null, null, null, sRutaPDF, null);
                                            }

                                        }
                                        else
                                        {

                                        }

                                        #endregion

                                        DllFunciones.sendMessageBox(_sboapp, "Codigo de mensaje No. " + RespuestaDoc.codigo.ToString() + ", " + "Documento autorizado por la DIAN");

                                        _sboapp.ActivateMenuItem("1304");

                                        #endregion

                                        DllFunciones.sendMessageBox(_sboapp, "Codigo de mensaje No. " + RespuestaDoc.codigo.ToString() + ", " + "Documentos autorizado por la DIAN");

                                        _sboapp.ActivateMenuItem("1304");

                                        #endregion
                                    }

                                }


                                #endregion
                            }
                            else
                            {
                                DllFunciones.sendMessageBox(_sboapp, "Error Paso 1: No se encontraron facturas para enviar");
                            }

                        }
                        #endregion

                        #endregion
                    }
                    else
                    {
                        #region Si el estado del documento es != 200, envia el documento a la DIAN

                        if (oConsultaDocEntry.RecordCount > 0)
                        {

                            DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 2: Creando Objeto " + sNombreDocWarning + " ...");

                            #region Si existe el numero de factura, busca la factura y crea el objeto factura

                            sDocEntryInvoice = oConsultaDocEntry.Fields.Item(0).Value.ToString();

                            SAPbobsCOM.Recordset oCabeceraDocumento = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            SAPbobsCOM.Recordset oLineasDocumento = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            SAPbobsCOM.Recordset oImpuestosGenerales = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            SAPbobsCOM.Recordset oImpuestosTotales = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            SAPbobsCOM.Recordset oCUFEInvoice = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                            sProcedureXML = DllFunciones.GetStringXMLDocument(_oCompany, "eBilling", "eBilling", "ExecProcedureBOFacturaXML");

                            if (_TipoDocumento == "FacturaDeClientes")
                            {
                                sDocumentoCabecera = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "13").Replace("%TipoConsulta%", "Encabezado");
                                sDocumentoLinea = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "13").Replace("%TipoConsulta%", "Lineas");
                                sDocumentoImpuestosGenerales = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "13").Replace("%TipoConsulta%", "Impuestos");
                                sDocumentoImpuestosTotales = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "13").Replace("%TipoConsulta%", "ImpuestosTotales");
                            }
                            else if (_TipoDocumento == "NotaCreditoClientes")
                            {
                                sDocumentoCabecera = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "14").Replace("%TipoConsulta%", "Encabezado");
                                sDocumentoLinea = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "14").Replace("%TipoConsulta%", "Lineas");
                                sDocumentoImpuestosGenerales = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "14").Replace("%TipoConsulta%", "Impuestos");
                                sDocumentoImpuestosTotales = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "14").Replace("%TipoConsulta%", "ImpuestosTotales");
                            }
                            else if (_TipoDocumento == "NotaDebitoClientes")
                            {
                                sDocumentoCabecera = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "13_ND").Replace("%TipoConsulta%", "Encabezado");
                                sDocumentoLinea = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "13_ND").Replace("%TipoConsulta%", "Lineas");
                                sDocumentoImpuestosGenerales = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "13_ND").Replace("%TipoConsulta%", "Impuestos");
                                sDocumentoImpuestosTotales = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "13_ND").Replace("%TipoConsulta%", "ImpuestosTotales");
                            }

                            oCabeceraDocumento.DoQuery(sDocumentoCabecera);
                            oLineasDocumento.DoQuery(sDocumentoLinea);
                            oImpuestosGenerales.DoQuery(sDocumentoImpuestosGenerales);
                            oImpuestosTotales.DoQuery(sDocumentoImpuestosTotales);

                            if (_TipoDocumento == "NotaCreditoClientes")
                            {
                                sCUFEInvoice = DllFunciones.GetStringXMLDocument(_oCompany, "eBilling", "eBilling", "GetCUFEInvoice");
                                sCUFEInvoice = sCUFEInvoice.Replace("%DocNum%", Convert.ToString(oCabeceraDocumento.Fields.Item("No_FV").Value.ToString()));

                                oCUFEInvoice.DoQuery(sCUFEInvoice);
                            }
                            else if (_TipoDocumento == "NotaDebitoClientes")
                            {
                                sCUFEInvoice = DllFunciones.GetStringXMLDocument(_oCompany, "eBilling", "eBilling", "GetCUFEDebitNote");
                                sCUFEInvoice = sCUFEInvoice.Replace("%DocNum%", Convert.ToString(oCabeceraDocumento.Fields.Item("No_FV").Value.ToString()));

                                oCUFEInvoice.DoQuery(sCUFEInvoice);

                            }
                            FacturaGeneral Documento = oBuillInvoice(oCabeceraDocumento, oLineasDocumento, oImpuestosGenerales, oImpuestosTotales, oCUFEInvoice, _TipoDocumento);

                            #endregion

                            DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 3: Guardando TXT " + sNombreDocWarning + " ...");

                            #region Guarda el TXT en la ruta del XML configurada

                            StreamWriter MyFile = new StreamWriter(sRutaXML); //ruta y name del archivo request a almecenar

                            #endregion

                            DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 4: Serealizando la " + sNombreDocWarning + " ...");

                            #region Serealizando el documento

                            SAPbobsCOM.Recordset oParametrosTFHKA = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                            sParametrosTFHKA = DllFunciones.GetStringXMLDocument(_oCompany, "eBilling", "eBilling", "GetParameterstoSend");

                            oParametrosTFHKA.DoQuery(sParametrosTFHKA);

                            XmlSerializer Serializer1 = new XmlSerializer(typeof(FacturaGeneral));
                            Serializer1.Serialize(MyFile, Documento); // Objeto serializado
                            MyFile.Close();

                            if (sGenerarXMLPrueba == "N")
                            {
                                File.Delete(sRutaXML);
                            }

                            #endregion

                            DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 5: Enviando " + sNombreDocWarning + " a TFHKA...");

                            #region Envio del objeto factura a TFHKA

                            serviceClient = new eBilling.ServicioEmisionFE.ServiceClient(port, endPointEmision);
                            serviceClientAdjuntos = new eBilling.ServicioAdjuntosFE.ServiceClient(port, endPointAdjuntos);

                            DocumentResponse RespuestaDoc = new eBilling.ServicioEmisionFE.DocumentResponse(); //objeto Response del metodo enviar

                            if (string.IsNullOrEmpty(sLlave))
                            {
                                DllFunciones.sendMessageBox(_sboapp, "Error Paso 4: No se ha parametrizado la llave de TFHKA en la configuracion Inicial, por lo cual no se puede enviar la factura a la DIAN ");
                            }
                            else if (string.IsNullOrEmpty(sPassword))
                            {
                                DllFunciones.sendMessageBox(_sboapp, "Error Paso 4: No se ha parametrizado el password de TFHKA en la configuracion Inicial, por lo cual no se puede enviar la factura a la DIAN");
                            }
                            else
                            {
                                #region Respuesta el Web Service de TFHKA y actualizacion de los campos en la factura

                                RespuestaDoc = serviceClient.Enviar(Convert.ToString(oParametrosTFHKA.Fields.Item("TokenEmpresa").Value.ToString()), Convert.ToString(oParametrosTFHKA.Fields.Item("TokenPassword").Value.ToString()), Documento, sFormaEnvio);

                                if (RespuestaDoc.codigo == 200)
                                {
                                    #region Procesa la repuesta

                                    DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 6: " + sNombreDocWarning + " enviada correctamente a TFHKA");

                                    #region Se actualiza el documento en SAP con las respuesta de TFHKA

                                    if (_TipoDocumento == "FacturaDeClientes" || (_TipoDocumento == "NotaDebitoClientes"))
                                    {
                                        UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, "Documento autorizado por la DIAN", RespuestaDoc.cufe, RespuestaDoc.qr, null, null);
                                        InsertSendEmail(_oCompany, oCabeceraDocumento, sCountsEmails, sDocEntryInvoice, "13");
                                    }
                                    else if (_TipoDocumento == "NotaCreditoClientes")
                                    {
                                        UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, "Documento autorizado por la DIAN", RespuestaDoc.cufe, RespuestaDoc.qr, null, null);
                                        InsertSendEmail(_oCompany, oCabeceraDocumento, sCountsEmails, sDocEntryInvoice, "14");
                                    }

                                    #endregion

                                    #region Valida la forma de envio,si es 11,  genera el PDF y retorna confirmacion de la generacion del PDF

                                    if (sFormaEnvio == "11")
                                    {
                                        DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 7: Generando PDF, por favor espere ...");

                                        FileInfo ValidacionPDF = new FileInfo(sRutaPDF);

                                        if (ValidacionPDF.Exists)
                                        {
                                            GeneroPDF = true;
                                        }
                                        else
                                        {
                                            GeneroPDF = ExportPDF(_sboapp, _oCompany, sRutaPDF, sDocEntryInvoice, sRutaCR, _TipoDocumento, sUserDB, sPassDB);
                                        }
                                    }

                                    #endregion

                                    #region Si genera correctamente el PDF lo adjunta a la factura de venta en SAP, 

                                    if (GeneroPDF == true)
                                    {
                                        DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 7: Adjuntando PDF, por favor espere ...");

                                        if (_TipoDocumento == "FacturaDeClientes")
                                        {
                                            UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, 0, null, null, null, sRutaPDF, null);
                                        }
                                        else if (_TipoDocumento == "NotaCreditoClientes")
                                        {
                                            UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, 0, null, null, null, sRutaPDF, null);
                                        }



                                    }
                                    else
                                    {
                                    }

                                    #endregion

                                    DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 8: Enviando PDF a TFHKA por favor espere ...");

                                    #region Envia el PDF al proveedor tecnologico TFHKA

                                    sPrefijoConDoc = Convert.ToString(oCabeceraDocumento.Fields.Item("Prefijo").Value.ToString()) + sDocNumInvoice;

                                    EnviarAdjuntosTFHKA(_sboapp, oCabeceraDocumento, sRutaPDF, sPrefijoConDoc, Convert.ToString(oParametrosTFHKA.Fields.Item("TokenEmpresa").Value.ToString()), Convert.ToString(oParametrosTFHKA.Fields.Item("TokenPassword").Value.ToString()));

                                    #endregion

                                    #region Se descarga el XML y se adjunta a la factura de venta

                                    DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 7: Adjuntando XML y actualizando campos, por favor espere ...");

                                    #region Descarga el XML y retorna la confirmacion

                                    bool DescargoXML = false;

                                    DescargoXML = DescargaXML(_oCompany, sPrefijoConDoc, Convert.ToString(oParametrosTFHKA.Fields.Item("TokenEmpresa").Value.ToString()), Convert.ToString(oParametrosTFHKA.Fields.Item("TokenPassword").Value.ToString()), sRutaXML);

                                    #endregion

                                    #region Actualiza el campo de XML en el documento de SAP


                                    if (DescargoXML == true)
                                    {
                                        UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, 0, null, null, null, null, sRutaXML.Replace(".txt", ".xml"));
                                    }
                                    else
                                    {

                                    }



                                    #endregion

                                    DllFunciones.sendMessageBox(_sboapp, "Codigo de mensaje No. " + RespuestaDoc.codigo.ToString() + ", " + "Documento enviado correctamente");

                                    _sboapp.ActivateMenuItem("1304");

                                    #endregion

                                    DllFunciones.sendMessageBox(_sboapp, "El documento fue enviado existosamente a la DIAN");

                                    DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 9: Proceso de emision Finalizado ...");

                                    _sboapp.ActivateMenuItem("1304");

                                    #endregion
                                }
                                else if (RespuestaDoc.codigo == 201)
                                {
                                    #region Procesa la respuesta                             

                                    DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 6: " + sNombreDocWarning + " enviada correctamente a TFHKA");

                                    #region Consulta el estado del documento en TFHKA

                                    sPrefijoConDoc = Convert.ToString(oCabeceraDocumento.Fields.Item("Prefijo").Value.ToString()) + sDocNumInvoice;
                                    DocumentStatusResponse resp = serviceClient.EstadoDocumento(Convert.ToString(oParametrosTFHKA.Fields.Item("TokenEmpresa").Value.ToString()), Convert.ToString(oParametrosTFHKA.Fields.Item("TokenPassword").Value.ToString()), sPrefijoConDoc);

                                    #endregion

                                    #region Actualiza el documento con la respuesta de TFHKA

                                    if (_TipoDocumento == "FacturaDeClientes" || (_TipoDocumento == "NotaDebitoClientes"))
                                    {
                                        UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, resp.codigo, "Documento autorizado por la DIAN", resp.cufe, resp.cadenaCodigoQR, null, null);
                                        InsertSendEmail(_oCompany, oCabeceraDocumento, sCountsEmails, sDocEntryInvoice, "13");
                                    }
                                    else if (_TipoDocumento == "NotaCreditoClientes")
                                    {
                                        UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, resp.codigo, "Documento autorizado por la DIAN", resp.cufe, resp.cadenaCodigoQR, null, null);
                                        InsertSendEmail(_oCompany, oCabeceraDocumento, sCountsEmails, sDocEntryInvoice, "14");
                                    }
                                    #endregion

                                    if (resp.codigo == 200)
                                    {
                                        #region Valida la forma de envio,si es 11,  genera el PDF y retorna confirmacion de la generacion del PDF

                                        if (sFormaEnvio == "11")
                                        {
                                            DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 7: Generando PDF, por favor espere ...");

                                            FileInfo ValidacionPDF = new FileInfo(sRutaPDF);

                                            if (ValidacionPDF.Exists)
                                            {
                                                GeneroPDF = true;
                                            }
                                            else
                                            {
                                                GeneroPDF = ExportPDF(_sboapp, _oCompany, sRutaPDF, sDocEntryInvoice, sRutaCR, _TipoDocumento, sUserDB, sPassDB);
                                            }

                                        }
                                        else
                                        {
                                            DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 8: Proceso de emision finalizado con exito.");
                                        }

                                        DllFunciones.sendMessageBox(_sboapp, "El documento fue enviado existosamente a la DIAN");

                                        #endregion
                                    }

                                    _sboapp.ActivateMenuItem("1304");

                                    #endregion
                                }
                                else if (RespuestaDoc.codigo == 101)
                                {
                                    #region Procesa la respuesta

                                    if (_TipoDocumento == "FacturaDeClientes" || (_TipoDocumento == "NotaDebitoClientes"))
                                    {
                                        UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, RespuestaDoc.mensaje, "", "", null, null);
                                    }
                                    else if (_TipoDocumento == "NotaCreditoClientes")
                                    {
                                        UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, RespuestaDoc.mensaje, "", "", null, null);
                                    }

                                    DllFunciones.sendMessageBox(_sboapp, "Codigo de error No. " + RespuestaDoc.codigo.ToString() + ", " + RespuestaDoc.mensaje.ToString());

                                    _sboapp.ActivateMenuItem("1304");

                                    #endregion
                                }
                                else if (RespuestaDoc.codigo == 99)
                                {
                                    #region Procesa la respuesta

                                    if (_TipoDocumento == "FacturaDeClientes" || (_TipoDocumento == "NotaDebitoClientes"))
                                    {
                                        UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, Convert.ToString(RespuestaDoc.reglasValidacionDIAN.GetValue(0)), "", "", null, null);
                                    }
                                    else if (_TipoDocumento == "NotaCreditoClientes")
                                    {
                                        UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, Convert.ToString(RespuestaDoc.reglasValidacionDIAN.GetValue(0)), "", "", null, null);
                                    }

                                    DllFunciones.sendMessageBox(_sboapp, "Codigo de error No. " + RespuestaDoc.codigo.ToString() + " , " + Convert.ToString(RespuestaDoc.reglasValidacionDIAN.GetValue(0)));

                                    _sboapp.ActivateMenuItem("1304");

                                    #endregion
                                }

                                else if (RespuestaDoc.codigo == 109)
                                {
                                    #region Procesa la respuesta

                                    if (_TipoDocumento == "FacturaDeClientes" || (_TipoDocumento == "NotaDebitoClientes"))
                                    {
                                        UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, Convert.ToString(RespuestaDoc.mensajesValidacion.GetValue(0)), "", "", null, null);
                                    }
                                    else if (_TipoDocumento == "NotaCreditoClientes")
                                    {
                                        UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, Convert.ToString(RespuestaDoc.mensajesValidacion.GetValue(0)), "", "", null, null);
                                    }

                                    DllFunciones.sendMessageBox(_sboapp, "Codigo de error No. " + RespuestaDoc.codigo.ToString() + ", " + RespuestaDoc.mensaje.ToString() + " " + RespuestaDoc.mensajesValidacion.GetValue(0));

                                    _sboapp.ActivateMenuItem("1304");

                                    #endregion
                                }
                                else if (RespuestaDoc.codigo == 110)
                                {
                                    #region Procesa la respuesta

                                    if (_TipoDocumento == "FacturaDeClientes" || (_TipoDocumento == "NotaDebitoClientes")) if (_TipoDocumento == "FacturaDeClientes")
                                        {
                                            UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, RespuestaDoc.mensaje.ToString() + ", Total de Factura es diferente de la suma de Total valor bruto + Tributos - Total Tributo Retenidos - Anticipos ", "", "", null, null);
                                        }
                                        else if (_TipoDocumento == "NotaCreditoClientes")
                                        {
                                            UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, RespuestaDoc.mensaje.ToString() + ", Total de Factura es diferente de la suma de Total valor bruto + Tributos - Total Tributo Retenidos - Anticipos ", "", "", null, null);
                                        }

                                    DllFunciones.sendMessageBox(_sboapp, "Codigo de error No. " + RespuestaDoc.codigo.ToString() + ", " + RespuestaDoc.mensaje.ToString() + ", Total de Factura es diferente de la suma de Total valor bruto + Tributos - Total Tributo Retenidos - Anticipos ");

                                    _sboapp.ActivateMenuItem("1304");

                                    #endregion
                                }
                                else if (RespuestaDoc.codigo == 111)
                                {
                                    #region Procesa la respuesta

                                    if (_TipoDocumento == "FacturaDeClientes" || (_TipoDocumento == "NotaDebitoClientes"))
                                    {
                                        UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, RespuestaDoc.mensaje.ToString(), "", "", null, null);
                                    }
                                    else if (_TipoDocumento == "NotaCreditoClientes")
                                    {
                                        UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, RespuestaDoc.mensaje.ToString(), "", "", null, null);
                                    }

                                    DllFunciones.sendMessageBox(_sboapp, "Codigo de error No. " + RespuestaDoc.codigo.ToString() + ", " + RespuestaDoc.mensaje.ToString());

                                    _sboapp.ActivateMenuItem("1304");

                                    #endregion

                                }
                                else if (RespuestaDoc.codigo == 112)
                                {
                                    #region Procesa la respuesta

                                    if (_TipoDocumento == "FacturaDeClientes" || (_TipoDocumento == "NotaDebitoClientes"))
                                    {
                                        UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, RespuestaDoc.mensaje.ToString(), "", "", null, null);
                                    }
                                    else if (_TipoDocumento == "NotaCreditoClientes")
                                    {
                                        UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, RespuestaDoc.mensaje.ToString(), "", "", null, null);
                                    }

                                    DllFunciones.sendMessageBox(_sboapp, "Codigo de error No. " + RespuestaDoc.codigo.ToString() + ", " + RespuestaDoc.mensaje.ToString());

                                    _sboapp.ActivateMenuItem("1304");

                                    #endregion
                                }
                                else if (RespuestaDoc.codigo == 150)
                                {
                                    #region Procesa la respuesta

                                    if (_TipoDocumento == "FacturaDeClientes" || (_TipoDocumento == "NotaDebitoClientes"))
                                    {
                                        UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, RespuestaDoc.mensaje.ToString(), "", "", null, null);
                                    }
                                    else if (_TipoDocumento == "NotaCreditoClientes")
                                    {
                                        UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, RespuestaDoc.mensaje.ToString(), "", "", null, null);
                                    }

                                    DllFunciones.sendMessageBox(_sboapp, "Codigo de error No. " + RespuestaDoc.codigo.ToString() + ", " + RespuestaDoc.mensaje.ToString());

                                    _sboapp.ActivateMenuItem("1304");

                                    #endregion
                                }
                                else if (RespuestaDoc.codigo == 114)
                                {
                                    #region Procesa la respuesta

                                    DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 5: Consultando el estado del documento a TFHKA, por favor espere ...");

                                    #region Consulta el estado del documento en el proveedor tecnologico

                                    sPrefijoConDoc = Convert.ToString(oCabeceraDocumento.Fields.Item("Prefijo").Value.ToString()) + sDocNumInvoice;
                                    DocumentStatusResponse resp = serviceClient.EstadoDocumento(Convert.ToString(oParametrosTFHKA.Fields.Item("TokenEmpresa").Value.ToString()), Convert.ToString(oParametrosTFHKA.Fields.Item("TokenPassword").Value.ToString()), sPrefijoConDoc);

                                    #endregion

                                    #region Se actualiza la factura con las respuesta de TFHKA

                                    if (_TipoDocumento == "FacturaDeClientes" || (_TipoDocumento == "NotaDebitoClientes"))
                                    {
                                        UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, resp.codigo, "Documento autorizado por la DIAN", resp.cufe, resp.cadenaCodigoQR, null, null);
                                        InsertSendEmail(_oCompany, oCabeceraDocumento, sCountsEmails, sDocEntryInvoice, "13");
                                    }
                                    else if (_TipoDocumento == "NotaCreditoClientes")
                                    {
                                        UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, resp.codigo, "Documento autorizado por la DIAN", resp.cufe, resp.cadenaCodigoQR, null, null);
                                        InsertSendEmail(_oCompany, oCabeceraDocumento, sCountsEmails, sDocEntryInvoice, "14");
                                    }

                                    #endregion

                                    #region Valida la forma de envio, si es 11 genera el PDF y retorna confirmacion de la generacion del PDF

                                    if (sFormaEnvio == "11")
                                    {
                                        DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 6: Generando PDF, por favor espere ...");

                                        FileInfo ValidacionPDF = new FileInfo(sRutaPDF);

                                        if (ValidacionPDF.Exists)
                                        {
                                            GeneroPDF = true;
                                        }
                                        else
                                        {
                                            GeneroPDF = ExportPDF(_sboapp, _oCompany, sRutaPDF, sDocEntryInvoice, sRutaCR, _TipoDocumento, sUserDB, sPassDB);
                                        }

                                    }

                                    #endregion

                                    #region Si genera correctamente el PDF lo adjunta a la factura de venta en SAP, 

                                    if (GeneroPDF == true)
                                    {
                                        DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 7: Adjuntando PDF y actualizando campos, por favor espere ...");

                                        if (_TipoDocumento == "FacturaDeClientes")
                                        {
                                            UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, resp.codigo, "Documento autorizado por la DIAN", resp.cufe, resp.cadenaCodigoQR, sRutaPDF, null);
                                        }
                                        else if (_TipoDocumento == "NotaCreditoClientes")
                                        {
                                            UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, resp.codigo, "Documento autorizado por la DIAN", resp.cufe, resp.cadenaCodigoQR, sRutaPDF, null);
                                        }

                                        #region Envia el PDF al proveedor tecnologico TFHKA

                                        DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 8: Enviando PDF a TFHKA por favor espere ...");

                                        EnviarAdjuntosTFHKA(_sboapp, oCabeceraDocumento, sRutaPDF, sPrefijoConDoc, sLlave, sPassword);

                                        #endregion

                                        DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 9: Proceso de emision Finalizado ...");

                                    }
                                    else
                                    {
                                        DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 7: Actualizando campos, por favor espere ...");

                                        UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, 0, null, null, null, null, null);

                                        DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 8: Proceso de emision Finalizado ...");
                                    }

                                    #endregion

                                    #region Se descarga el XML y se adjunta a la factura de venta

                                    DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 7: Adjuntando XML y actualizando campos, por favor espere ...");

                                    #region Descarga el XML y retorna la confirmacion

                                    bool DescargoXML = false;

                                    DescargoXML = DescargaXML(_oCompany, sPrefijoConDoc, Convert.ToString(oParametrosTFHKA.Fields.Item("TokenEmpresa").Value.ToString()), Convert.ToString(oParametrosTFHKA.Fields.Item("TokenPassword").Value.ToString()), sRutaXML);

                                    #endregion

                                    #region Actualiza el campo de XML en el documento de SAP

                                    if (DescargoXML == true)
                                    {

                                        if (_TipoDocumento == "FacturaDeClientes")
                                        {
                                            UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, 0, null, null, null, null, sRutaXML.Replace(".txt", ".xml"));

                                        }
                                        else if (_TipoDocumento == "NotaCreditoClientes")
                                        {
                                            UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, 0, null, null, null, sRutaPDF, null);
                                        }

                                    }
                                    else
                                    {

                                    }

                                    #endregion

                                    DllFunciones.sendMessageBox(_sboapp, "Codigo de mensaje No. " + RespuestaDoc.codigo.ToString() + ", " + "Documento autorizado por la DIAN");

                                    _sboapp.ActivateMenuItem("1304");

                                    #endregion

                                    DllFunciones.sendMessageBox(_sboapp, "Codigo de mensaje No. " + RespuestaDoc.codigo.ToString() + ", " + "Documentos autorizado por la DIAN");

                                    _sboapp.ActivateMenuItem("1304");

                                    #endregion
                                }
                                else if (RespuestaDoc.codigo == 1)
                                {
                                    #region Procesa la respuesta

                                    DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 5: Consultando el estado del documento a TFHKA, por favor espere ...");

                                    #region Consulta el estado del documento en el proveedor tecnologico

                                    sPrefijoConDoc = Convert.ToString(oCabeceraDocumento.Fields.Item("Prefijo").Value.ToString()) + sDocNumInvoice;
                                    DocumentStatusResponse resp = serviceClient.EstadoDocumento(Convert.ToString(oParametrosTFHKA.Fields.Item("TokenEmpresa").Value.ToString()), Convert.ToString(oParametrosTFHKA.Fields.Item("TokenPassword").Value.ToString()), sPrefijoConDoc);

                                    #endregion

                                    #region Se actualiza la factura con las respuesta de TFHKA

                                    if (_TipoDocumento == "FacturaDeClientes" || (_TipoDocumento == "NotaDebitoClientes"))
                                    {
                                        UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, resp.codigo, "Documento autorizado por la DIAN", resp.cufe, resp.cadenaCodigoQR, null, null);
                                        InsertSendEmail(_oCompany, oCabeceraDocumento, sCountsEmails, sDocEntryInvoice, "13");
                                    }
                                    else if (_TipoDocumento == "NotaCreditoClientes")
                                    {
                                        UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, resp.codigo, "Documento autorizado por la DIAN", resp.cufe, resp.cadenaCodigoQR, null, null);
                                        InsertSendEmail(_oCompany, oCabeceraDocumento, sCountsEmails, sDocEntryInvoice, "14");
                                    }

                                    #endregion

                                    #region Valida la forma de envio, si es 11 genera el PDF y retorna confirmacion de la generacion del PDF

                                    if (sFormaEnvio == "11")
                                    {
                                        DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 6: Generando PDF, por favor espere ...");

                                        FileInfo ValidacionPDF = new FileInfo(sRutaPDF);

                                        if (ValidacionPDF.Exists)
                                        {
                                            GeneroPDF = true;
                                        }
                                        else
                                        {
                                            GeneroPDF = ExportPDF(_sboapp, _oCompany, sRutaPDF, sDocEntryInvoice, sRutaCR, _TipoDocumento, sUserDB, sPassDB);
                                        }

                                    }

                                    #endregion

                                    #region Si genera correctamente el PDF lo adjunta a la factura de venta en SAP, 

                                    if (GeneroPDF == true)
                                    {
                                        DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 7: Adjuntando PDF y actualizando campos, por favor espere ...");

                                        if (_TipoDocumento == "FacturaDeClientes")
                                        {
                                            UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, resp.codigo, "Documento autorizado por la DIAN", resp.cufe, resp.cadenaCodigoQR, sRutaPDF, null);
                                        }
                                        else if (_TipoDocumento == "NotaCreditoClientes")
                                        {
                                            UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, resp.codigo, "Documento autorizado por la DIAN", resp.cufe, resp.cadenaCodigoQR, sRutaPDF, null);
                                        }

                                        #region Envia el PDF al proveedor tecnologico TFHKA

                                        DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 8: Enviando PDF a TFHKA por favor espere ...");

                                        EnviarAdjuntosTFHKA(_sboapp, oCabeceraDocumento, sRutaPDF, sPrefijoConDoc, sLlave, sPassword);

                                        #endregion

                                        DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 9: Proceso de emision Finalizado ...");

                                    }
                                    else
                                    {
                                        DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 7: Actualizando campos, por favor espere ...");

                                        UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, 0, null, null, null, null, null);

                                        DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 8: Proceso de emision Finalizado ...");
                                    }

                                    #endregion

                                    #region Se descarga el XML y se adjunta a la factura de venta

                                    DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 7: Adjuntando XML y actualizando campos, por favor espere ...");

                                    #region Descarga el XML y retorna la confirmacion

                                    bool DescargoXML = false;

                                    DescargoXML = DescargaXML(_oCompany, sPrefijoConDoc, Convert.ToString(oParametrosTFHKA.Fields.Item("TokenEmpresa").Value.ToString()), Convert.ToString(oParametrosTFHKA.Fields.Item("TokenPassword").Value.ToString()), sRutaXML);

                                    #endregion

                                    #region Actualiza el campo de XML en el documento de SAP

                                    if (DescargoXML == true)
                                    {

                                        if (_TipoDocumento == "FacturaDeClientes")
                                        {
                                            UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, 0, null, null, null, null, sRutaXML.Replace(".txt", ".xml"));

                                        }
                                        else if (_TipoDocumento == "NotaCreditoClientes")
                                        {
                                            UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, 0, null, null, null, sRutaPDF, null);
                                        }

                                    }
                                    else
                                    {

                                    }

                                    #endregion

                                    DllFunciones.sendMessageBox(_sboapp, "Codigo de mensaje No. " + RespuestaDoc.codigo.ToString() + ", " + "Documento autorizado por la DIAN");

                                    _sboapp.ActivateMenuItem("1304");

                                    #endregion

                                    DllFunciones.sendMessageBox(_sboapp, "Codigo de mensaje No. " + RespuestaDoc.codigo.ToString() + ", " + "Documentos autorizado por la DIAN");

                                    _sboapp.ActivateMenuItem("1304");

                                    #endregion
                                }

                                #endregion
                            }

                            #endregion
                        }
                        else
                        {
                            DllFunciones.sendMessageBox(_sboapp, "Error Paso 1: No se encontraron facturas para enviar");
                        }

                        #endregion
                    }

                    #endregion

                }
            }
        }

        public void EnviarDocumentosMasivamenteTFHKA(SAPbouiCOM.Application _sboapp, SAPbobsCOM.Company _oCompany, SAPbouiCOM.Form _oFormDocuments, string _TipoDocumento, string TipoIntegracion, int iNumeroLinea)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            if (TipoIntegracion == "M")
            {
                #region Envio del documento por el visor de documentos

                #region Consulta URL

                string sGetModo = null;
                string sURLEmision = null;
                string sURLAdjuntos = null;
                string sModo = null;

                SAPbobsCOM.Recordset oConsultarGetModo = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                sGetModo = DllFunciones.GetStringXMLDocument(_oCompany, "eBilling", "eBilling", "GetModoandURL");

                sGetModo = sGetModo.Replace("%Estado%", "\"U_BO_Status\" = 'Y'").Replace("%DocEntry%", " ");

                oConsultarGetModo.DoQuery(sGetModo);

                sURLEmision = Convert.ToString(oConsultarGetModo.Fields.Item("URLTFHKA").Value.ToString()) + "/ws/v1.0/Service.svc?wsdl";
                sURLAdjuntos = Convert.ToString(oConsultarGetModo.Fields.Item("URLTFHKA").Value.ToString()) + "/ws/adjuntos/Service.svc?wsdl";
                sModo = Convert.ToString(oConsultarGetModo.Fields.Item("Modo").Value.ToString());

                DllFunciones.liberarObjetos(oConsultarGetModo);

                #endregion

                #region Instanciacion parametros TFHKA

                //Especifica el puerto (HTTP o HTTPS)
                if (sModo == "PRU")
                {
                    BasicHttpBinding port = new BasicHttpBinding();
                }
                else if (sModo == "PRO")
                {
                    BasicHttpsBinding port = new BasicHttpsBinding();
                }

                port.MaxBufferPoolSize = Int32.MaxValue;
                port.MaxBufferSize = Int32.MaxValue;
                port.MaxReceivedMessageSize = Int32.MaxValue;
                port.ReaderQuotas.MaxStringContentLength = Int32.MaxValue;
                port.SendTimeout = TimeSpan.FromMinutes(2);
                port.ReceiveTimeout = TimeSpan.FromMinutes(2);

                if (sModo == "PRO")
                {
                    port.Security.Mode = BasicHttpSecurityMode.Transport;
                }

                //Especifica la dirección de conexion para Emision y Adjuntos 
                EndpointAddress endPointEmision = new EndpointAddress(sURLEmision); //URL DEMO EMISION
                EndpointAddress endPointAdjuntos = new EndpointAddress(sURLAdjuntos); //URL DEMO ADJUNTOS          

                #endregion

                #region Variables y objetos

                string sDocNumInvoice = null;
                string sQueryDocEntryDocument = null;
                string sDocEntryInvoice = null;
                string sProcedureXML = null;
                string sDocumentoCabecera = null;
                string sDocumentoLinea = null;
                string sDocumentoImpuestosGenerales = null;
                string sDocumentoImpuestosTotales = null;
                string sParametrosTFHKA = null;
                string sRutaCR = null;
                string sPrefijoConDoc = null;
                string sPrefijo = null;
                string sStatusDoc = null;
                string sFormaEnvio = null;
                string sLlave = null;
                string sPassword = null;
                string sUserDB = null;
                string sPassDB = null;
                string sRutaPDF = null;
                string sRutaXML = null;
                string sNombreDocumento = null;
                string sNombreDocWarning = null;
                string sCUFEInvoice = null;
                int sReprocesar = 0;
                string sGenerarXMLPrueba = null;

                Boolean GeneroPDF = false;

                if (_TipoDocumento == "FacturaDeClientes")
                {
                    sNombreDocumento = "Factura_de_Venta_No_";
                    sNombreDocWarning = "Factura de venta";
                }
                else if (_TipoDocumento == "NotaCreditoClientes")
                {
                    sNombreDocumento = "Nota_Credito_No_";
                    sNombreDocWarning = "Nota credito de clientes";
                }
                else if (_TipoDocumento == "NotaDebitoClientes")
                {
                    sNombreDocumento = "Nota_debito_Clientes_No_";
                    sNombreDocWarning = "Nota debito de clientes";
                }

                SAPbouiCOM.Matrix oMatrixInvoice = (Matrix)_oFormDocuments.Items.Item("MtxOINV").Specific;
                SAPbouiCOM.Matrix oMatrixCreditMemo = (Matrix)_oFormDocuments.Items.Item("MtxORIN").Specific;
                SAPbouiCOM.Matrix oMatrixDebitMemo = (Matrix)_oFormDocuments.Items.Item("MtxOINVD").Specific;


                #endregion

                #region Consulta de documento en la base de datos y el estado del documento

                sDocEntryInvoice = ((SAPbouiCOM.EditText)(oMatrixInvoice.Columns.Item("Col_9").Cells.Item(iNumeroLinea).Specific)).Value;

                SAPbobsCOM.Recordset oConsultaDocEntry = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                sQueryDocEntryDocument = DllFunciones.GetStringXMLDocument(_oCompany, "eBilling", "eBilling", "GetDocEntryAndParametersService");

                if (_TipoDocumento == "FacturaDeClientes")
                {
                    sQueryDocEntryDocument = sQueryDocEntryDocument.Replace("%sDocEntry%", sDocEntryInvoice).Replace("%Tabla%", "OINV").Replace("%DocSubType%", "--");
                }
                else if (_TipoDocumento == "NotaCreditoClientes")
                {
                    sQueryDocEntryDocument = sQueryDocEntryDocument.Replace("%sDocEntry%", sDocEntryInvoice).Replace("%Tabla%", "ORIN").Replace("%DocSubType%", "--");
                }
                else if (_TipoDocumento == "NotaDebitoClientes")
                {
                    sQueryDocEntryDocument = sQueryDocEntryDocument.Replace("%sDocEntry%", sDocEntryInvoice).Replace("%Tabla%", "OINV").Replace("%DocSubType%", "DN");
                }

                oConsultaDocEntry.DoQuery(sQueryDocEntryDocument);

                sDocNumInvoice = Convert.ToString(oConsultaDocEntry.Fields.Item("DocNum").Value.ToString());
                sPrefijo = Convert.ToString(oConsultaDocEntry.Fields.Item("PrefijoDes").Value.ToString());
                sStatusDoc = Convert.ToString(oConsultaDocEntry.Fields.Item("CRWS").Value.ToString());
                sFormaEnvio = Convert.ToString(oConsultaDocEntry.Fields.Item("FormaEnvio").Value.ToString());
                sLlave = Convert.ToString(oConsultaDocEntry.Fields.Item("Llave").Value.ToString());
                sPassword = Convert.ToString(oConsultaDocEntry.Fields.Item("Password").Value.ToString());
                sUserDB = Convert.ToString(oConsultaDocEntry.Fields.Item("UserDB").Value.ToString());
                sPassDB = Convert.ToString(oConsultaDocEntry.Fields.Item("PassDB").Value.ToString());
                sRutaXML = Convert.ToString(oConsultaDocEntry.Fields.Item("RutaXML").Value.ToString()) + "\\" + sNombreDocumento + sPrefijo + '_' + sDocNumInvoice + ".txt";
                sRutaPDF = Convert.ToString(oConsultaDocEntry.Fields.Item("RutaPDF").Value.ToString()) + "\\" + sNombreDocumento + sPrefijo + '_' + sDocNumInvoice + ".pdf";
                sRutaCR = Convert.ToString(oConsultaDocEntry.Fields.Item("RutaCR").Value.ToString());
                sGenerarXMLPrueba = Convert.ToString(oConsultaDocEntry.Fields.Item("GeneraXMLP").Value.ToString());

                #endregion

                if (sStatusDoc == "200")
                {
                    DllFunciones.liberarObjetos(oMatrixInvoice);
                    DllFunciones.liberarObjetos(oMatrixCreditMemo);
                    DllFunciones.liberarObjetos(oMatrixDebitMemo);
                    DllFunciones.liberarObjetos(oConsultaDocEntry);
                    DllFunciones.liberarObjetos(oConsultarGetModo);
                }
                else
                {
                    if (oConsultaDocEntry.RecordCount > 0)
                    {

                        #region Si existe el numero de factura, busca la factura y crea el objeto factura

                        SAPbobsCOM.Recordset oCabeceraDocumento = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        SAPbobsCOM.Recordset oLineasDocumento = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        SAPbobsCOM.Recordset oImpuestosGenerales = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        SAPbobsCOM.Recordset oImpuestosTotales = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        SAPbobsCOM.Recordset oCUFEInvoice = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                        sProcedureXML = DllFunciones.GetStringXMLDocument(_oCompany, "eBilling", "eBilling", "ExecProcedureBOFacturaXML");

                        if (_TipoDocumento == "FacturaDeClientes")
                        {
                            sDocumentoCabecera = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "13").Replace("%TipoConsulta%", "Encabezado");
                            sDocumentoLinea = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "13").Replace("%TipoConsulta%", "Lineas");
                            sDocumentoImpuestosGenerales = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "13").Replace("%TipoConsulta%", "Impuestos");
                            sDocumentoImpuestosTotales = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "13").Replace("%TipoConsulta%", "ImpuestosTotales");
                        }
                        else if (_TipoDocumento == "NotaCreditoClientes")
                        {
                            sDocumentoCabecera = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "14").Replace("%TipoConsulta%", "Encabezado");
                            sDocumentoLinea = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "14").Replace("%TipoConsulta%", "Lineas");
                            sDocumentoImpuestosGenerales = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "14").Replace("%TipoConsulta%", "Impuestos");
                            sDocumentoImpuestosTotales = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "14").Replace("%TipoConsulta%", "ImpuestosTotales");
                        }
                        else if (_TipoDocumento == "NotaDebitoClientes")
                        {
                            sDocumentoCabecera = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "13_ND").Replace("%TipoConsulta%", "Encabezado");
                            sDocumentoLinea = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "13_ND").Replace("%TipoConsulta%", "Lineas");
                            sDocumentoImpuestosGenerales = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "13_ND").Replace("%TipoConsulta%", "Impuestos");
                            sDocumentoImpuestosTotales = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "13_ND").Replace("%TipoConsulta%", "ImpuestosTotales");
                        }

                        oCabeceraDocumento.DoQuery(sDocumentoCabecera);
                        oLineasDocumento.DoQuery(sDocumentoLinea);
                        oImpuestosGenerales.DoQuery(sDocumentoImpuestosGenerales);
                        oImpuestosTotales.DoQuery(sDocumentoImpuestosTotales);

                        if (_TipoDocumento == "NotaCreditoClientes")
                        {
                            sCUFEInvoice = DllFunciones.GetStringXMLDocument(_oCompany, "eBilling", "eBilling", "GetCUFEInvoice");
                            sCUFEInvoice = sCUFEInvoice.Replace("%DocNum%", Convert.ToString(oCabeceraDocumento.Fields.Item("No_FV").Value.ToString()));

                            oCUFEInvoice.DoQuery(sCUFEInvoice);
                        }
                        else if (_TipoDocumento == "NotaDebitoClientes")
                        {
                            sCUFEInvoice = DllFunciones.GetStringXMLDocument(_oCompany, "eBilling", "eBilling", "GetCUFEDebitNote");
                            sCUFEInvoice = sCUFEInvoice.Replace("%DocNum%", Convert.ToString(oCabeceraDocumento.Fields.Item("No_FV").Value.ToString()));

                            oCUFEInvoice.DoQuery(sCUFEInvoice);

                        }
                        FacturaGeneral Documento = oBuillInvoice(oCabeceraDocumento, oLineasDocumento, oImpuestosGenerales, oImpuestosTotales, oCUFEInvoice, _TipoDocumento);

                        #endregion

                        #region Guarda el TXT en la ruta del XML configurada

                        StreamWriter MyFile = new StreamWriter(sRutaXML); //ruta y name del archivo request a almecenar

                        #endregion

                        #region Serealizando el documento

                        SAPbobsCOM.Recordset oParametrosTFHKA = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                        sParametrosTFHKA = DllFunciones.GetStringXMLDocument(_oCompany, "eBilling", "eBilling", "GetParameterstoSend");

                        oParametrosTFHKA.DoQuery(sParametrosTFHKA);

                        XmlSerializer Serializer1 = new XmlSerializer(typeof(FacturaGeneral));
                        Serializer1.Serialize(MyFile, Documento); // Objeto serializado
                        MyFile.Close();

                        if (sGenerarXMLPrueba == "N")
                        {
                            File.Delete(sRutaXML);
                        }

                        #endregion

                        #region Envio del objeto factura a TFHKA

                        serviceClient = new eBilling.ServicioEmisionFE.ServiceClient(port, endPointEmision);
                        serviceClientAdjuntos = new eBilling.ServicioAdjuntosFE.ServiceClient(port, endPointAdjuntos);

                        DocumentResponse RespuestaDoc = new eBilling.ServicioEmisionFE.DocumentResponse(); //objeto Response del metodo enviar

                        if (string.IsNullOrEmpty(sLlave))
                        {
                            DllFunciones.sendMessageBox(_sboapp, "Error Paso 4: No se ha parametrizado la llave de TFHKA en la configuracion Inicial, por lo cual no se puede enviar la factura a la DIAN ");
                        }
                        else if (string.IsNullOrEmpty(sPassword))
                        {
                            DllFunciones.sendMessageBox(_sboapp, "Error Paso 4: No se ha parametrizado el password de TFHKA en la configuracion Inicial, por lo cual no se puede enviar la factura a la DIAN");
                        }
                        else
                        {
                            #region Respuesta el Web Service de TFHKA y actualizacion de los campos en la factura

                            RespuestaDoc = serviceClient.Enviar(Convert.ToString(oParametrosTFHKA.Fields.Item("TokenEmpresa").Value.ToString()), Convert.ToString(oParametrosTFHKA.Fields.Item("TokenPassword").Value.ToString()), Documento, sFormaEnvio);

                            if (RespuestaDoc.codigo == 200)
                            {
                                #region Procesa la repuesta

                                #region Se actualiza el documento en SAP con las respuesta de TFHKA

                                if (_TipoDocumento == "FacturaDeClientes" || (_TipoDocumento == "NotaDebitoClientes"))
                                {
                                    UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, Convert.ToString(RespuestaDoc.mensaje), RespuestaDoc.cufe, RespuestaDoc.qr, null, null);
                                }
                                else if (_TipoDocumento == "NotaCreditoClientes")
                                {
                                    UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, Convert.ToString(RespuestaDoc.mensaje), RespuestaDoc.cufe, RespuestaDoc.qr, null, null);
                                }

                                #endregion

                                #region Valida la forma de envio,si es 11,  genera el PDF y retorna confirmacion de la generacion del PDF

                                if (sFormaEnvio == "11")
                                {
                                    //DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 7: Generando PDF, por favor espere ...");

                                    FileInfo ValidacionPDF = new FileInfo(sRutaPDF);

                                    if (ValidacionPDF.Exists)
                                    {
                                        GeneroPDF = true;
                                    }
                                    else
                                    {
                                        GeneroPDF = ExportPDF(_sboapp, _oCompany, sRutaPDF, sDocEntryInvoice, sRutaCR, _TipoDocumento, sUserDB, sPassDB);
                                    }
                                }

                                #endregion

                                #region Si genera correctamente el PDF lo adjunta a la factura de venta en SAP, 

                                if (GeneroPDF == true)
                                {
                                    //DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 7: Adjuntando PDF, por favor espere ...");

                                    if (_TipoDocumento == "FacturaDeClientes")
                                    {
                                        UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, 0, null, null, null, sRutaPDF, null);
                                    }
                                    else if (_TipoDocumento == "NotaCreditoClientes")
                                    {
                                        UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, 0, null, null, null, sRutaPDF, null);
                                    }

                                }
                                else
                                {
                                }

                                #endregion

                                #region Envia el PDF al proveedor tecnologico TFHKA

                                sPrefijoConDoc = Convert.ToString(oCabeceraDocumento.Fields.Item("Prefijo").Value.ToString()) + sDocNumInvoice;

                                EnviarAdjuntosTFHKA(_sboapp, oCabeceraDocumento, sRutaPDF, sPrefijoConDoc, Convert.ToString(oParametrosTFHKA.Fields.Item("TokenEmpresa").Value.ToString()), Convert.ToString(oParametrosTFHKA.Fields.Item("TokenPassword").Value.ToString()));

                                #endregion

                                #region Se descarga el XML y se adjunta a la factura de venta   

                                #region Descarga el XML y retorna la confirmacion

                                bool DescargoXML = false;

                                DescargoXML = DescargaXML(_oCompany, sPrefijoConDoc, Convert.ToString(oParametrosTFHKA.Fields.Item("TokenEmpresa").Value.ToString()), Convert.ToString(oParametrosTFHKA.Fields.Item("TokenPassword").Value.ToString()), sRutaXML);

                                #endregion

                                #region Actualiza el campo de XML en el documento de SAP

                                if (DescargoXML == true)
                                {
                                    UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, 0, null, null, null, null, sRutaXML.Replace(".txt", ".xml"));
                                }
                                else
                                {

                                }

                                #endregion

                                #endregion

                                #endregion
                            }
                            else if (RespuestaDoc.codigo == 201)
                            {
                                #region Procesa la respuesta                                                        

                                #region Consulta el estado del documento en TFHKA

                                sPrefijoConDoc = Convert.ToString(oCabeceraDocumento.Fields.Item("Prefijo").Value.ToString()) + sDocNumInvoice;
                                DocumentStatusResponse resp = serviceClient.EstadoDocumento(Convert.ToString(oParametrosTFHKA.Fields.Item("TokenEmpresa").Value.ToString()), Convert.ToString(oParametrosTFHKA.Fields.Item("TokenPassword").Value.ToString()), sPrefijoConDoc);

                                #endregion

                                #region Actualiza el documento con la respuesta de TFHKA

                                if (_TipoDocumento == "FacturaDeClientes" || (_TipoDocumento == "NotaDebitoClientes"))
                                {
                                    UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, resp.codigo, "Procesado Correctamente", resp.cufe, resp.cadenaCodigoQR, null, null);
                                }
                                else if (_TipoDocumento == "NotaCreditoClientes")
                                {
                                    UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, resp.codigo, "Procesado Correctamente", resp.cufe, resp.cadenaCodigoQR, null, null);
                                }
                                #endregion

                                if (resp.codigo == 200)
                                {
                                    #region Valida la forma de envio,si es 11,  genera el PDF y retorna confirmacion de la generacion del PDF

                                    if (sFormaEnvio == "11")
                                    {
                                        FileInfo ValidacionPDF = new FileInfo(sRutaPDF);

                                        if (ValidacionPDF.Exists)
                                        {
                                            GeneroPDF = true;
                                        }
                                        else
                                        {
                                            GeneroPDF = ExportPDF(_sboapp, _oCompany, sRutaPDF, sDocEntryInvoice, sRutaCR, _TipoDocumento, sUserDB, sPassDB);
                                        }

                                    }
                                    else
                                    {

                                    }

                                    #endregion
                                }

                                #endregion
                            }
                            else if (RespuestaDoc.codigo == 101)
                            {
                                #region Procesa la respuesta

                                if (_TipoDocumento == "FacturaDeClientes" || (_TipoDocumento == "NotaDebitoClientes"))
                                {
                                    UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, RespuestaDoc.mensaje, "", "", null, null);
                                }
                                else if (_TipoDocumento == "NotaCreditoClientes")
                                {
                                    UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, RespuestaDoc.mensaje, "", "", null, null);
                                }

                                DllFunciones.sendMessageBox(_sboapp, "Codigo de error No. " + RespuestaDoc.codigo.ToString() + ", " + RespuestaDoc.mensaje.ToString());

                                #endregion
                            }
                            else if (RespuestaDoc.codigo == 99)
                            {
                                #region Procesa la respuesta

                                if (_TipoDocumento == "FacturaDeClientes" || (_TipoDocumento == "NotaDebitoClientes"))
                                {
                                    UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, Convert.ToString(RespuestaDoc.reglasValidacionDIAN.GetValue(0)), "", "", null, null);
                                }
                                else if (_TipoDocumento == "NotaCreditoClientes")
                                {
                                    UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, Convert.ToString(RespuestaDoc.mensajesValidacion.GetValue(0)), "", "", null, null);
                                }

                                #endregion
                            }

                            else if (RespuestaDoc.codigo == 109)
                            {
                                #region Procesa la respuesta

                                if (_TipoDocumento == "FacturaDeClientes" || (_TipoDocumento == "NotaDebitoClientes"))
                                {
                                    UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, Convert.ToString(RespuestaDoc.mensajesValidacion.GetValue(0)), "", "", null, null);
                                }
                                else if (_TipoDocumento == "NotaCreditoClientes")
                                {
                                    UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, Convert.ToString(RespuestaDoc.mensajesValidacion.GetValue(0)), "", "", null, null);
                                }

                                #endregion
                            }
                            else if (RespuestaDoc.codigo == 110)
                            {
                                #region Procesa la respuesta

                                if (_TipoDocumento == "FacturaDeClientes" || (_TipoDocumento == "NotaDebitoClientes")) if (_TipoDocumento == "FacturaDeClientes")
                                    {
                                        UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, RespuestaDoc.mensaje.ToString() + ", Total de Factura es diferente de la suma de Total valor bruto + Tributos - Total Tributo Retenidos - Anticipos ", "", "", null, null);
                                    }
                                    else if (_TipoDocumento == "NotaCreditoClientes")
                                    {
                                        UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, RespuestaDoc.mensaje.ToString() + ", Total de Factura es diferente de la suma de Total valor bruto + Tributos - Total Tributo Retenidos - Anticipos ", "", "", null, null);
                                    }

                                #endregion
                            }
                            else if (RespuestaDoc.codigo == 111)
                            {
                                #region Procesa la respuesta

                                if (_TipoDocumento == "FacturaDeClientes" || (_TipoDocumento == "NotaDebitoClientes"))
                                {
                                    UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, RespuestaDoc.mensaje.ToString(), "", "", null, null);
                                }
                                else if (_TipoDocumento == "NotaCreditoClientes")
                                {
                                    UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, RespuestaDoc.mensaje.ToString(), "", "", null, null);
                                }

                                #endregion

                            }
                            else if (RespuestaDoc.codigo == 112)
                            {
                                #region Procesa la respuesta

                                if (_TipoDocumento == "FacturaDeClientes" || (_TipoDocumento == "NotaDebitoClientes"))
                                {
                                    UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, RespuestaDoc.mensaje.ToString(), "", "", null, null);
                                }
                                else if (_TipoDocumento == "NotaCreditoClientes")
                                {
                                    UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, RespuestaDoc.mensaje.ToString(), "", "", null, null);
                                }

                                #endregion
                            }
                            else if (RespuestaDoc.codigo == 150)
                            {
                                #region Procesa la respuesta

                                if (_TipoDocumento == "FacturaDeClientes" || (_TipoDocumento == "NotaDebitoClientes"))
                                {
                                    UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, RespuestaDoc.mensaje.ToString(), "", "", null, null);
                                }
                                else if (_TipoDocumento == "NotaCreditoClientes")
                                {
                                    UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, RespuestaDoc.codigo, RespuestaDoc.mensaje.ToString(), "", "", null, null);
                                }

                                #endregion
                            }
                            else if (RespuestaDoc.codigo == 114)
                            {
                                #region Procesa la respuesta

                                #region Consulta el estado del documento en el proveedor tecnologico

                                sPrefijoConDoc = Convert.ToString(oCabeceraDocumento.Fields.Item("Prefijo").Value.ToString()) + sDocNumInvoice;
                                DocumentStatusResponse resp = serviceClient.EstadoDocumento(Convert.ToString(oParametrosTFHKA.Fields.Item("TokenEmpresa").Value.ToString()), Convert.ToString(oParametrosTFHKA.Fields.Item("TokenPassword").Value.ToString()), sPrefijoConDoc);

                                #endregion

                                #region Se actualiza la factura con las respuesta de TFHKA

                                if (_TipoDocumento == "FacturaDeClientes" || (_TipoDocumento == "NotaDebitoClientes"))
                                {
                                    UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, resp.codigo, resp.mensajeDocumento, resp.cufe, resp.cadenaCodigoQR, null, null);
                                }
                                else if (_TipoDocumento == "NotaCreditoClientes")
                                {
                                    UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, resp.codigo, resp.mensajeDocumento, resp.cufe, resp.cadenaCodigoQR, null, null);
                                }

                                #endregion

                                #region Valida la forma de envio, si es 11 genera el PDF y retorna confirmacion de la generacion del PDF

                                if (sFormaEnvio == "11")
                                {
                                    FileInfo ValidacionPDF = new FileInfo(sRutaPDF);

                                    if (ValidacionPDF.Exists)
                                    {
                                        GeneroPDF = true;
                                    }
                                    else
                                    {
                                        GeneroPDF = ExportPDF(_sboapp, _oCompany, sRutaPDF, sDocEntryInvoice, sRutaCR, _TipoDocumento, sUserDB, sPassDB);
                                    }

                                }

                                #endregion

                                #region Si genera correctamente el PDF lo adjunta a la factura de venta en SAP, 

                                if (GeneroPDF == true)
                                {
                                    if (_TipoDocumento == "FacturaDeClientes")
                                    {
                                        UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, resp.codigo, resp.mensajeDocumento, resp.cufe, resp.cadenaCodigoQR, sRutaPDF, null);
                                    }
                                    else if (_TipoDocumento == "NotaCreditoClientes")
                                    {
                                        UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, resp.codigo, resp.mensajeDocumento, resp.cufe, resp.cadenaCodigoQR, sRutaPDF, null);
                                    }

                                    #region Envia el PDF al proveedor tecnologico TFHKA

                                    EnviarAdjuntosTFHKA(_sboapp, oCabeceraDocumento, sRutaPDF, sPrefijoConDoc, sLlave, sPassword);

                                    #endregion

                                }
                                else
                                {
                                    DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 7: Actualizando campos, por favor espere ...");

                                    UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, 0, null, null, null, null, null);

                                    DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 8: Proceso de emision Finalizado ...");
                                }

                                #endregion

                                #region Se descarga el XML y se adjunta a la factura de venta

                                #region Descarga el XML y retorna la confirmacion

                                bool DescargoXML = false;

                                DescargoXML = DescargaXML(_oCompany, sPrefijoConDoc, Convert.ToString(oParametrosTFHKA.Fields.Item("TokenEmpresa").Value.ToString()), Convert.ToString(oParametrosTFHKA.Fields.Item("TokenPassword").Value.ToString()), sRutaXML);

                                #endregion

                                #region Actualiza el campo de XML en el documento de SAP

                                if (DescargoXML == true)
                                {

                                    if (_TipoDocumento == "FacturaDeClientes")
                                    {
                                        UpdateoInvoice(_oCompany, _sboapp, sDocEntryInvoice, 0, null, null, null, null, sRutaXML.Replace(".txt", ".xml"));

                                    }
                                    else if (_TipoDocumento == "NotaCreditoClientes")
                                    {
                                        UpdateoCreditNote(_oCompany, _sboapp, sDocEntryInvoice, 0, null, null, null, sRutaPDF, null);
                                    }

                                }
                                else
                                {

                                }

                                #endregion

                                #endregion

                                #endregion
                            }

                            #endregion
                        }

                        #endregion
                    }
                    else
                    {

                    }
                    #endregion
                }
            }
            else if (TipoIntegracion == "S")
            {


            }
        }

        public void EnviarAdjuntosTFHKA(SAPbouiCOM.Application _sboapp, SAPbobsCOM.Recordset oCabecera, string _RutaPDFyXML, string _sPrefijoConDoc, string _tbxTokenEmpresa, string _tbxTokenPassword)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            int procesados = 0;

            for (int i = 0; i < 1; i++)
            {
                FileInfo file = new FileInfo(_RutaPDFyXML);

                if (file.Exists)
                {
                    BinaryReader bReader = new BinaryReader(file.OpenRead());
                    byte[] anexByte = bReader.ReadBytes((int)file.Length);
                    //anexB64 = Convert.ToBase64String(anexByte);
                    ServicioAdjuntosFE.CargarAdjuntos uploadAttachment = new ServicioAdjuntosFE.CargarAdjuntos();
                    uploadAttachment.archivo = anexByte;
                    uploadAttachment.numeroDocumento = _sPrefijoConDoc;

                    #region Revision Correos a Enviar

                    #region Variables Correo

                    string CorreoDeEntrega1 = Convert.ToString(oCabecera.Fields.Item("correoEntrega1").Value.ToString());
                    string CorreoDeEntrega2 = Convert.ToString(oCabecera.Fields.Item("correoEntrega2").Value.ToString());
                    string CorreoDeEntrega3 = Convert.ToString(oCabecera.Fields.Item("correoEntrega3").Value.ToString());
                    string CorreoDeEntrega4 = Convert.ToString(oCabecera.Fields.Item("correoEntrega4").Value.ToString());
                    string CorreoDeEntrega5 = Convert.ToString(oCabecera.Fields.Item("correoEntrega5").Value.ToString());

                    int ContadorCorreos = 0;

                    #endregion

                    #region Contador de los correos a enviar 

                    if (string.IsNullOrEmpty(CorreoDeEntrega1))
                    {

                    }
                    else
                    {
                        ContadorCorreos++;
                    }

                    if (string.IsNullOrEmpty(CorreoDeEntrega2))
                    {

                    }
                    else
                    {
                        ContadorCorreos++;
                    }

                    if (string.IsNullOrEmpty(CorreoDeEntrega3))
                    {

                    }
                    else
                    {
                        ContadorCorreos++;
                    }

                    if (string.IsNullOrEmpty(CorreoDeEntrega4))
                    {

                    }
                    else
                    {
                        ContadorCorreos++;
                    }

                    if (string.IsNullOrEmpty(CorreoDeEntrega5))
                    {

                    }
                    else
                    {
                        ContadorCorreos++;
                    }

                    #endregion

                    string[] correoEntrega = new string[ContadorCorreos];

                    #region Asignacion de los correos a enviar

                    if (ContadorCorreos == 1)
                    {
                        correoEntrega[0] = CorreoDeEntrega1;
                    }
                    else if (ContadorCorreos == 2)
                    {
                        correoEntrega[0] = CorreoDeEntrega1;
                        correoEntrega[1] = CorreoDeEntrega2;
                    }
                    else if (ContadorCorreos == 3)
                    {
                        correoEntrega[0] = CorreoDeEntrega1;
                        correoEntrega[1] = CorreoDeEntrega2;
                        correoEntrega[2] = CorreoDeEntrega3;
                    }
                    else if (ContadorCorreos == 4)
                    {
                        correoEntrega[0] = CorreoDeEntrega1;
                        correoEntrega[1] = CorreoDeEntrega2;
                        correoEntrega[2] = CorreoDeEntrega3;
                        correoEntrega[3] = CorreoDeEntrega4;
                    }
                    else if (ContadorCorreos == 5)
                    {
                        correoEntrega[0] = CorreoDeEntrega1;
                        correoEntrega[1] = CorreoDeEntrega2;
                        correoEntrega[2] = CorreoDeEntrega3;
                        correoEntrega[3] = CorreoDeEntrega4;
                        correoEntrega[4] = CorreoDeEntrega5;
                    }

                    #endregion

                    #endregion

                    uploadAttachment.email = correoEntrega;
                    uploadAttachment.nombre = file.Name.Substring(0, file.Name.Length - 4);
                    uploadAttachment.formato = file.Extension.Substring(1);
                    uploadAttachment.tipo = "1";

                    if (Convert.ToString(oCabecera.Fields.Item("notificar").Value.ToString()) == "NO")
                    {
                        uploadAttachment.enviar = "0";
                    }
                    else
                    {
                        if (i + 1 == 1)
                        {
                            uploadAttachment.enviar = "1";
                        }
                        else
                        {
                            uploadAttachment.enviar = "0";
                        }
                    }


                    ServicioAdjuntosFE.UploadAttachmentResponse fileRespuesta = new ServicioAdjuntosFE.UploadAttachmentResponse();
                    fileRespuesta = serviceClientAdjuntos.CargarAdjuntos(_tbxTokenEmpresa, _tbxTokenPassword, uploadAttachment);
                    if (fileRespuesta.codigo == 200)
                    {

                    }
                    else
                    {
                        DllFunciones.sendMessageBox(_sboapp, " No se cargo correctamente el PDF al portal de TFHKA, " + fileRespuesta.mensaje);
                    }
                }
                else
                {
                    // no debería entrar a este ciclo
                }

            }

        }

        private void UpdateoInvoice(SAPbobsCOM.Company __oCompany, SAPbouiCOM.Application __sboapp, string _sQueryDocEntryInvoice, int _CRWS, string _MRWS, string _WSCUFE, string _WSQR, string _RutaPDF, string _RutaXML)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            try
            {
                int _DocEntry = Convert.ToInt32(_sQueryDocEntryInvoice);
                Rsd = 0;

                SAPbobsCOM.Documents oInvoice = (SAPbobsCOM.Documents)(__oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices));

                oInvoice.GetByKey(_DocEntry);

                #region Campo CRWS

                if (_CRWS == 0)
                {
                }
                else
                {
                    oInvoice.UserFields.Fields.Item("U_BO_CRWS").Value = Convert.ToString(_CRWS);
                }

                #endregion

                #region Campo MRWS

                if (string.IsNullOrEmpty(_MRWS))
                {
                }
                else
                {
                    oInvoice.UserFields.Fields.Item("U_BO_MRWS").Value = Convert.ToString(_MRWS);
                }

                #endregion

                oInvoice.UserFields.Fields.Item("U_BO_S").Value = "3";
                oInvoice.UserFields.Fields.Item("U_BO_PP").Value = "A";

                #region Campo WSCUFE

                if (string.IsNullOrEmpty(_WSCUFE))
                {
                }
                else
                {
                    oInvoice.UserFields.Fields.Item("U_BO_CUFE").Value = Convert.ToString(_WSCUFE);
                }

                #endregion

                #region Campo WSQR

                if (string.IsNullOrEmpty(_WSQR))
                {
                }
                else
                {
                    oInvoice.UserFields.Fields.Item("U_BO_QR").Value = Convert.ToString(_WSQR);
                }

                #endregion

                #region Campo RutaPDF

                if (string.IsNullOrEmpty(_RutaPDF))
                {
                }
                else
                {
                    oInvoice.UserFields.Fields.Item("U_BO_RPDF").Value = Convert.ToString(_RutaPDF);
                }

                #endregion

                #region Campo RutaXML

                if (string.IsNullOrEmpty(_RutaXML))
                {
                }
                else
                {
                    oInvoice.UserFields.Fields.Item("U_BO_XML").Value = Convert.ToString(_RutaXML);
                }

                #endregion

                Rsd = oInvoice.Update();

                if (Rsd == 0)
                {
                    DllFunciones.liberarObjetos(oInvoice);
                }
                else
                {
                    DllFunciones.sendMessageBox(__sboapp, __oCompany.GetLastErrorDescription());
                }
            }
            catch (Exception e)
            {
                DllFunciones.sendErrorMessage(__sboapp, e);
            }

        }

        private void UpdateoCreditNote(SAPbobsCOM.Company __oCompany, SAPbouiCOM.Application __sboapp, string _sQueryDocEntryInvoice, int _CRWS, string _MRWS, string _WSCUFE, string _WSQR, string _RutaPDF, string _RutaXML)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            try
            {
                int _DocEntry = Convert.ToInt32(_sQueryDocEntryInvoice);
                Rsd = 0;

                SAPbobsCOM.Documents oCreditNote = (SAPbobsCOM.Documents)(__oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes));

                oCreditNote.GetByKey(_DocEntry);

                #region Campo CRWS

                if (_CRWS == 0)
                {
                }
                else
                {
                    oCreditNote.UserFields.Fields.Item("U_BO_CRWS").Value = Convert.ToString(_CRWS);
                }

                #endregion

                #region Campo MRWS

                if (string.IsNullOrEmpty(_MRWS))
                {
                }
                else
                {
                    oCreditNote.UserFields.Fields.Item("U_BO_MRWS").Value = Convert.ToString(_MRWS);
                }

                #endregion

                oCreditNote.UserFields.Fields.Item("U_BO_S").Value = "3";
                oCreditNote.UserFields.Fields.Item("U_BO_PP").Value = "A";

                #region Campo WSCUFE

                if (string.IsNullOrEmpty(_WSCUFE))
                {
                }
                else
                {
                    oCreditNote.UserFields.Fields.Item("U_BO_CUFE").Value = Convert.ToString(_WSCUFE);
                }

                #endregion

                #region Campo WSQR

                if (string.IsNullOrEmpty(_WSQR))
                {
                }
                else
                {
                    oCreditNote.UserFields.Fields.Item("U_BO_QR").Value = Convert.ToString(_WSQR);
                }

                #endregion

                #region Campo RutaPDF

                if (string.IsNullOrEmpty(_RutaPDF))
                {
                }
                else
                {
                    oCreditNote.UserFields.Fields.Item("U_BO_RPDF").Value = Convert.ToString(_RutaPDF);
                }

                #endregion

                #region Campo RutaXML

                if (string.IsNullOrEmpty(_RutaXML))
                {
                }
                else
                {
                    oCreditNote.UserFields.Fields.Item("U_BO_XML").Value = Convert.ToString(_RutaXML);
                }

                #endregion

                Rsd = oCreditNote.Update();

                if (Rsd == 0)
                {
                    DllFunciones.liberarObjetos(oCreditNote);
                }
                else
                {
                    DllFunciones.sendMessageBox(__sboapp, __oCompany.GetLastErrorDescription());
                }
            }
            catch (Exception e)
            {
                DllFunciones.sendErrorMessage(__sboapp, e);
            }

        }

        private bool DescargaXML(SAPbobsCOM.Company _oCompany, string _sPrefijoConDoc, string _tbxTokenEmpresa, string _tbxTokenPassword, string _sRutaXML)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            try
            {
                #region Consulta URL

                string sGetModo = null;
                string sURLEmision = null;
                string sURLAdjuntos = null;
                string sModo = null;
                string sRutaXML = null;

                SAPbobsCOM.Recordset oConsultarGetModo = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                sGetModo = DllFunciones.GetStringXMLDocument(_oCompany, "eBilling", "eBilling", "GetModoandURL");

                sGetModo = sGetModo.Replace("%Estado%", "\"U_BO_Status\" = 'Y'").Replace("%DocEntry%", " ");

                oConsultarGetModo.DoQuery(sGetModo);

                sURLEmision = Convert.ToString(oConsultarGetModo.Fields.Item("URLTFHKA").Value.ToString()) + "/ws/v1.0/Service.svc?wsdl";
                sURLAdjuntos = Convert.ToString(oConsultarGetModo.Fields.Item("URLTFHKA").Value.ToString()) + "/ws/adjuntos/Service.svc?wsdl";
                sModo = Convert.ToString(oConsultarGetModo.Fields.Item("Modo").Value.ToString());
                sRutaXML = _sRutaXML.Replace(".txt", ".xml");


                DllFunciones.liberarObjetos(oConsultarGetModo);

                #endregion

                #region Instanciacion parametros TFHKA

                //Especifica el puerto (HTTP o HTTPS)
                if (sModo == "PRU")
                {
                    BasicHttpBinding port = new BasicHttpBinding();
                }
                else if (sModo == "PRO")
                {
                    BasicHttpsBinding port = new BasicHttpsBinding();
                }

                port.MaxBufferPoolSize = Int32.MaxValue;
                port.MaxBufferSize = Int32.MaxValue;
                port.MaxReceivedMessageSize = Int32.MaxValue;
                port.ReaderQuotas.MaxStringContentLength = Int32.MaxValue;
                port.SendTimeout = TimeSpan.FromMinutes(2);
                port.ReceiveTimeout = TimeSpan.FromMinutes(2);

                if (sModo == "PRO")
                {
                    port.Security.Mode = BasicHttpSecurityMode.Transport;
                }

                //Especifica la dirección de conexion para Demo y Adjuntos para pruebas
                EndpointAddress endPointEmision = new EndpointAddress(sURLEmision); //URL DEMO EMISION

                ServicioEmisionFE.ServiceClient serviceClienTFHKA;

                serviceClienTFHKA = new ServicioEmisionFE.ServiceClient(port, endPointEmision);

                #endregion

                DownloadXMLResponse xmlResponse;

                xmlResponse = serviceClient.DescargaXML(_tbxTokenEmpresa, _tbxTokenPassword, _sPrefijoConDoc);

                if (xmlResponse.codigo == 200)
                {
                    File.WriteAllBytes(sRutaXML, Convert.FromBase64String(xmlResponse.documento));
                    return true;
                }
                else
                {
                    return false;
                }


            }
            catch (Exception)
            {

                return false;
            }

        }

        public void ChooFormListSN(string _FormUID, SAPbouiCOM.Form _FormVD, ItemEvent pVal, SAPbobsCOM.Company _oCompany, SAPbouiCOM.Application _sboapp)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            #region Variables y Objetos

            string sCFL_ID = null;
            string _sMotorDB = (Convert.ToString(_oCompany.DbServerType));

            SAPbouiCOM.EditText otxtSN = (SAPbouiCOM.EditText)_FormVD.Items.Item("txtSN").Specific;

            SAPbouiCOM.IChooseFromListEvent oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));

            sCFL_ID = oCFLEvento.ChooseFromListUID;

            SAPbouiCOM.Form _oFormRC = _sboapp.Forms.Item(_FormUID);

            SAPbouiCOM.ChooseFromList oCFL = _oFormRC.ChooseFromLists.Item(sCFL_ID);

            #endregion

            if (oCFLEvento.BeforeAction == false)
            {

                #region Variables y Objetos 

                string val = null;

                SAPbouiCOM.DataTable oDataTable = oCFLEvento.SelectedObjects;

                #endregion

                try
                {
                    val = System.Convert.ToString(oDataTable.GetValue(0, 0));

                    if (_sMotorDB == "dst_MSSQL2012" || _sMotorDB == "dst_MSSQL2014" || _sMotorDB == "dst_MSSQL2016" || _sMotorDB == "dst_HANADB")
                    {
                        if (pVal.ItemUID == "txtSN")
                        {
                            _oFormRC.DataSources.UserDataSources.Item("EditDS").ValueEx = val;
                        }
                    }
                }
                catch (Exception e)
                {
                    DllFunciones.sendErrorMessage(_sboapp, e);
                }

                if ((_FormUID == "CFL1") & (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD))
                {
                    System.Windows.Forms.Application.Exit();
                }
            }
        }

        private void AddChooseFromList(SAPbouiCOM.Application _sboapp, SAPbouiCOM.Form _oFormVD)
        {

            SAPbouiCOM.ChooseFromListCollection oCFLs = null;
            SAPbouiCOM.Conditions oCons = null;
            SAPbouiCOM.Condition oCon = null;

            oCFLs = _oFormVD.ChooseFromLists;

            SAPbouiCOM.ChooseFromList oCFL = null;
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
            oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(_sboapp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)));

            //  Adding 2 CFL, one for the button and one for the edit text.
            oCFLCreationParams.MultiSelection = false;
            oCFLCreationParams.ObjectType = "2";
            oCFLCreationParams.UniqueID = "CFL1";

            oCFL = oCFLs.Add(oCFLCreationParams);

            //  Adding Conditions to CFL1

            oCons = oCFL.GetConditions();

            oCon = oCons.Add();
            oCon.Alias = "CardType";
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCon.CondVal = "C";
            oCFL.SetConditions(oCons);

            oCFLCreationParams.UniqueID = "CFL2";
            oCFL = oCFLs.Add(oCFLCreationParams);
        }

        public void Right_Click(ref SAPbouiCOM.ContextMenuInfo _eventInfo, SAPbouiCOM.Application _sboapp, string NumeroID)
        {
            SAPbouiCOM.MenuItem oMenuItem = null;
            SAPbouiCOM.Menus oMenus = null;

            try
            {

                if (NumeroID == "1")
                {
                    #region Click derecho para adicionar linea en Matrix en Series numeracion

                    if (_sboapp.Menus.Exists("AddRowMtx"))
                    {

                    }
                    else
                    {
                        SAPbouiCOM.MenuCreationParams oCreationPackage = null;
                        oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(_sboapp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));

                        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                        oCreationPackage.UniqueID = "AddRowMtx";
                        oCreationPackage.String = "Añadir Serie Numeración";
                        oCreationPackage.Enabled = true;

                        oMenuItem = _sboapp.Menus.Item("1280"); // Data'
                        oMenus = oMenuItem.SubMenus;
                        oMenus.AddEx(oCreationPackage);

                    }

                    #endregion
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void AddRowMatrix(SAPbouiCOM.Form _oFormPareBilling)
        {
            SAPbouiCOM.Matrix oMatrixSeres = (Matrix)_oFormPareBilling.Items.Item("MtxSN").Specific;

            oMatrixSeres.AddRow();

        }

        public void InsertDataSeriesNumber(SAPbobsCOM.Company _oCompany, SAPbouiCOM.Form _oFormParametros)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            #region Variables y objetos

            SAPbobsCOM.GeneralService oGeneralService;
            SAPbobsCOM.GeneralData oGeneralData;
            SAPbobsCOM.CompanyService oCS = _oCompany.GetCompanyService();

            SAPbouiCOM.Matrix oMatrixSN = (Matrix)_oFormParametros.Items.Item("MtxSN").Specific;

            SAPbobsCOM.Recordset oConsultaCode = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string sGetCodeOriginal = null;
            string sGetCodeFinal = null;
            int iContador = 0;

            #endregion

            iContador = oMatrixSN.RowCount;

            if (iContador > 0)
            {
                oGeneralService = oCS.GetGeneralService("BOSERNUM");

                oGeneralData = (SAPbobsCOM.GeneralData)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);

                int iLinea = 1;

                sGetCodeOriginal = DllFunciones.GetStringXMLDocument(_oCompany, "eBilling", "eBilling", "GetCodeSeriesNumeracion");

                for (int i = 1; i <= iContador; i++)
                {


                    sGetCodeFinal = sGetCodeOriginal.Replace("%Code%", ((SAPbouiCOM.EditText)oMatrixSN.Columns.Item("Col_02").Cells.Item(iLinea).Specific).Value);

                    oConsultaCode.DoQuery(sGetCodeFinal);

                    if (oConsultaCode.RecordCount > 0)
                    {

                        #region Si no existe, inserta la serie de numeracion

                        oGeneralData.SetProperty("Code", ((SAPbouiCOM.EditText)oMatrixSN.Columns.Item("Col_02").Cells.Item(iLinea).Specific).Value);
                        oGeneralData.SetProperty("U_BO_SN", ((SAPbouiCOM.EditText)oMatrixSN.Columns.Item("Col_02").Cells.Item(iLinea).Specific).Value);
                        oGeneralData.SetProperty("U_BO_NR", ((SAPbouiCOM.EditText)oMatrixSN.Columns.Item("Col_03").Cells.Item(iLinea).Specific).Value);

                        string Fecha = ((SAPbouiCOM.EditText)oMatrixSN.Columns.Item("Col_04").Cells.Item(iLinea).Specific).Value;

                        if (string.IsNullOrEmpty(Fecha))
                        {

                        }
                        else
                        {
                            Fecha = Fecha.Insert(4, "-").Insert(7, "-");
                            oGeneralData.SetProperty("U_BO_FR", Fecha);
                        }

                        oGeneralData.SetProperty("U_BO_PREF", ((SAPbouiCOM.EditText)oMatrixSN.Columns.Item("Col_05").Cells.Item(iLinea).Specific).Value);
                        oGeneralData.SetProperty("U_BO_TD", ((SAPbouiCOM.ComboBox)oMatrixSN.Columns.Item("Col_01").Cells.Item(iLinea).Specific).Value);

                        oGeneralService.Update(oGeneralData);

                        #endregion

                    }
                    else
                    {
                        #region Si existe la serie de numeracion solamente la actualiza

                        string a = ((SAPbouiCOM.EditText)oMatrixSN.Columns.Item("Col_02").Cells.Item(iLinea).Specific).Value;

                        oGeneralData.SetProperty("Code", ((SAPbouiCOM.EditText)oMatrixSN.Columns.Item("Col_02").Cells.Item(iLinea).Specific).Value);
                        oGeneralData.SetProperty("U_BO_SN", ((SAPbouiCOM.EditText)oMatrixSN.Columns.Item("Col_02").Cells.Item(iLinea).Specific).Value);
                        oGeneralData.SetProperty("U_BO_NR", ((SAPbouiCOM.EditText)oMatrixSN.Columns.Item("Col_03").Cells.Item(iLinea).Specific).Value);
                        oGeneralData.SetProperty("U_BO_FR", ((SAPbouiCOM.EditText)oMatrixSN.Columns.Item("Col_04").Cells.Item(iLinea).Specific).Value);
                        oGeneralData.SetProperty("U_BO_PREF", ((SAPbouiCOM.EditText)oMatrixSN.Columns.Item("Col_05").Cells.Item(iLinea).Specific).Value);
                        oGeneralData.SetProperty("U_BO_TD", ((SAPbouiCOM.ComboBox)oMatrixSN.Columns.Item("Col_01").Cells.Item(iLinea).Specific).Value);

                        oGeneralService.Add(oGeneralData);

                        #endregion
                    }

                    iLinea++;

                }
            }
        }

        public void ActualizarEstadoDocumentos(SAPbouiCOM.Application _sboapp, SAPbobsCOM.Company oCompany, SAPbouiCOM.Form oFormVD)
        {

            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            #region Variables y Objetos

            int ContadorInvoice = 0;
            int iProcesar = 0;

            SAPbouiCOM.Matrix oMatrixInvoice = (Matrix)oFormVD.Items.Item("MtxOINV").Specific;
            SAPbouiCOM.Matrix oMatrixCreditMemo = (Matrix)oFormVD.Items.Item("MtxORIN").Specific;
            SAPbouiCOM.Matrix oMatrixDebitMemo = (Matrix)oFormVD.Items.Item("MtxOINVD").Specific;

            #endregion

            ContadorInvoice = oMatrixInvoice.RowCount + oMatrixCreditMemo.RowCount + oMatrixDebitMemo.RowCount;

            if (ContadorInvoice > 0)
            {
                iProcesar = DllFunciones.sendMessageBoxY_N(_sboapp, "Se enviara a la DIAN los documentos pendientes, esta seguro ? ");

                if (iProcesar == 1)
                {
                    #region Enviando Facturas de Venta

                    for (int i = 1; i <= oMatrixInvoice.RowCount; i++)
                    {
                        DllFunciones.ProgressBar(oCompany, _sboapp, ContadorInvoice, 1, "Enviando la factura de venta No. " + ((SAPbouiCOM.EditText)(oMatrixInvoice.Columns.Item("Col_1").Cells.Item(i).Specific)).Value);

                        EnviarDocumentosMasivamenteTFHKA(sboapp, oCompany, oFormVD, "FacturaDeClientes", "M", i);
                    }

                    #endregion

                    #region Enviando Notas Credito de Venta

                    for (int j = 1; j <= oMatrixCreditMemo.RowCount; j++)
                    {
                        DllFunciones.ProgressBar(oCompany, _sboapp, ContadorInvoice, 1, "Enviando la Nota Credito de Venta No. " + ((SAPbouiCOM.EditText)(oMatrixInvoice.Columns.Item("Col_1").Cells.Item(j).Specific)).Value);

                        EnviarDocumentosMasivamenteTFHKA(sboapp, oCompany, oFormVD, "NotaCreditoClientes", "M", j);
                    }

                    #endregion

                    #region Enviando Notas Debito de Cliente

                    for (int k = 1; k <= oMatrixDebitMemo.RowCount; k++)
                    {
                        DllFunciones.ProgressBar(oCompany, _sboapp, ContadorInvoice, 1, "Enviando la Nota Debito No. " + ((SAPbouiCOM.EditText)(oMatrixInvoice.Columns.Item("Col_1").Cells.Item(k).Specific)).Value);

                        EnviarDocumentosMasivamenteTFHKA(sboapp, oCompany, oFormVD, "NotaDebitoClientes", "M", k);
                    }

                    #endregion

                    DllFunciones.sendMessageBox(_sboapp, "Todos los documentos fueron enviados correctamente");
                }
            }
        }

        public void InsertSendEmail(SAPbobsCOM.Company _oCompany, SAPbobsCOM.Recordset _oCabecera, string _sCountsEmails, string _sDocEntry, string _sObjecType)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();
            try
            {

                if (Convert.ToString(_oCabecera.Fields.Item("notificar").Value.ToString()) == "SI")
                {
                    #region Variables y objetos

                    SAPbobsCOM.Recordset oConsultaDoc = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    #endregion

                    #region Consulta si ya se guardo el correo en la tablas de correos

                    _sCountsEmails = _sCountsEmails.Replace("%DocEntry%", _sDocEntry).Replace("%ObjecType%", _sObjecType);

                    oConsultaDoc.DoQuery(_sCountsEmails);

                    #endregion

                    if (oConsultaDoc.RecordCount > 0)
                    {

                    }
                    else
                    {
                        #region Inserta el correo en la tablas de correos

                        #region Variables y objetos

                        string _sSerachNextCode;

                        SAPbobsCOM.Recordset oSerachNextCode = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                        _sSerachNextCode = DllFunciones.GetStringXMLDocument(_oCompany, "eBilling", "eBilling", "SerachNextCode");

                        oSerachNextCode.DoQuery(_sSerachNextCode);

                        #endregion

                        #region Asignacion de valores

                        SAPbobsCOM.UserTable oUserTable;

                        oUserTable = _oCompany.UserTables.Item("BOEE");
                        oUserTable.Code = Convert.ToString(oSerachNextCode.Fields.Item("ID").Value.ToString());
                        oUserTable.Name = Convert.ToString(oSerachNextCode.Fields.Item("ID").Value.ToString());
                        oUserTable.UserFields.Fields.Item("U_BO_DocEntry").Value = _sDocEntry;
                        oUserTable.UserFields.Fields.Item("U_BO_ObjecType").Value = _sObjecType;

                        #region Asignacion Correo 1

                        if (string.IsNullOrWhiteSpace(Convert.ToString(_oCabecera.Fields.Item("correoEntrega1").Value.ToString())))
                        {

                        }
                        else
                        {
                            oUserTable.UserFields.Fields.Item("U_BO_Email1").Value = Convert.ToString(_oCabecera.Fields.Item("correoEntrega1").Value.ToString());
                        }

                        #endregion

                        #region Asignacion Correo 2

                        if (string.IsNullOrWhiteSpace(Convert.ToString(_oCabecera.Fields.Item("correoEntrega2").Value.ToString())))
                        {

                        }
                        else
                        {
                            oUserTable.UserFields.Fields.Item("U_BO_Email2").Value = Convert.ToString(_oCabecera.Fields.Item("correoEntrega2").Value.ToString());
                        }

                        #endregion

                        #region Asignacion Correo 3

                        if (string.IsNullOrWhiteSpace(Convert.ToString(_oCabecera.Fields.Item("correoEntrega3").Value.ToString())))
                        {

                        }
                        else
                        {
                            oUserTable.UserFields.Fields.Item("U_BO_Email3").Value = Convert.ToString(_oCabecera.Fields.Item("correoEntrega3").Value.ToString());
                        }

                        #endregion

                        #region Asignacion Correo 4

                        if (string.IsNullOrWhiteSpace(Convert.ToString(_oCabecera.Fields.Item("correoEntrega4").Value.ToString())))
                        {

                        }
                        else
                        {
                            oUserTable.UserFields.Fields.Item("U_BO_Email4").Value = Convert.ToString(_oCabecera.Fields.Item("correoEntrega4").Value.ToString());
                        }

                        #endregion

                        #region Asignacion Correo 5

                        if (string.IsNullOrWhiteSpace(Convert.ToString(_oCabecera.Fields.Item("correoEntrega5").Value.ToString())))
                        {

                        }
                        else
                        {
                            oUserTable.UserFields.Fields.Item("U_BO_Email5").Value = Convert.ToString(_oCabecera.Fields.Item("correoEntrega5").Value.ToString());
                        }

                        #endregion

                        #endregion

                        oUserTable.Add();

                        #endregion

                        DllFunciones.liberarObjetos(oSerachNextCode);

                    }

                    DllFunciones.liberarObjetos(oConsultaDoc);
                }
            }
            catch (Exception)
            {

                throw;
            }


        }

        public void EnviarCorreo(SAPbobsCOM.Company _oCompany, SAPbouiCOM.Application _sboapp, string sPrefijoDocumentoSM)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            try
            {
                #region StatusBar Enviando correo

                DllFunciones.sendStatusBarMsg(_sboapp, "Enviando correo electronico, por favor espere....", BoMessageTime.bmt_Short, false);

                #endregion

                #region Consulta URL

                string sGetModo = null;
                string sURLEmision = null;
                string sURLAdjuntos = null;
                string sModo = null;

                SAPbobsCOM.Recordset oConsultarGetModo = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                sGetModo = DllFunciones.GetStringXMLDocument(_oCompany, "eBilling", "eBilling", "GetModoandURL");

                sGetModo = sGetModo.Replace("%Estado%", "\"U_BO_Status\" = 'Y'").Replace("%DocEntry%", " ");

                oConsultarGetModo.DoQuery(sGetModo);

                sURLEmision = Convert.ToString(oConsultarGetModo.Fields.Item("URLTFHKA").Value.ToString()) + "/ws/v1.0/Service.svc?wsdl";
                sModo = Convert.ToString(oConsultarGetModo.Fields.Item("Modo").Value.ToString());

                DllFunciones.liberarObjetos(oConsultarGetModo);

                #endregion

                #region Instanciacion parametros TFHKA

                //Especifica el puerto (HTTP o HTTPS)
                if (sModo == "PRU")
                {
                    BasicHttpBinding port = new BasicHttpBinding();
                }
                else if (sModo == "PRO")
                {
                    BasicHttpsBinding port = new BasicHttpsBinding();
                }

                port.MaxBufferPoolSize = Int32.MaxValue;
                port.MaxBufferSize = Int32.MaxValue;
                port.MaxReceivedMessageSize = Int32.MaxValue;
                port.ReaderQuotas.MaxStringContentLength = Int32.MaxValue;
                port.SendTimeout = TimeSpan.FromMinutes(2);
                port.ReceiveTimeout = TimeSpan.FromMinutes(2);

                if (sModo == "PRO")
                {
                    port.Security.Mode = BasicHttpSecurityMode.Transport;
                }

                //Especifica la dirección de conexion para Emision y Adjuntos 
                EndpointAddress endPointEmision = new EndpointAddress(sURLEmision); //URL DEMO EMISION      

                #endregion

                #region Variables

                serviceClient = new eBilling.ServicioEmisionFE.ServiceClient(port, endPointEmision);

                string sCadenaCorreos = null;
                string sCorreoPrefijo = null;

                #endregion

                #region obtiene formulario correo 

                SAPbouiCOM.Form oFormSM;
                oFormSM = _sboapp.Forms.Item("BO_SM");

                SAPbouiCOM.Button btnCancel = (SAPbouiCOM.Button)(oFormSM.Items.Item("btnClose").Specific);

                #endregion

                #region Consulta y obtinene y el password

                SAPbobsCOM.Recordset oLlaveyPassword = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string sQueryDocEntryDocument = DllFunciones.GetStringXMLDocument(_oCompany, "eBilling", "eBilling", "GetLlaveAndPassword");

                oLlaveyPassword.DoQuery(sQueryDocEntryDocument);

                #endregion

                #region Consultar cantidad de correos 

                string sQuantityEmails;
                int iQuantityEmails;

                SAPbobsCOM.Recordset oQuantityEmails = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                sQuantityEmails = DllFunciones.GetStringXMLDocument(_oCompany, "eBilling", "eBilling", "GetQuantityEmails");
                oQuantityEmails.DoQuery(sQuantityEmails);

                iQuantityEmails = Convert.ToInt32(oQuantityEmails.Fields.Item(0).Value.ToString());

                DllFunciones.liberarObjetos(oQuantityEmails);

                #endregion

                #region Se obtiene el numero de documento

                SAPbouiCOM.StaticText olblL1 = (SAPbouiCOM.StaticText)(oFormSM.Items.Item("lbl1").Specific);

                #endregion

                #region Envio del correo

                if (iQuantityEmails == 2)
                {
                    #region Obtiene la cadena de correos

                    SAPbouiCOM.EditText txtEmail1 = (SAPbouiCOM.EditText)(oFormSM.Items.Item("txtEmail1").Specific);
                    SAPbouiCOM.EditText txtEmail2 = (SAPbouiCOM.EditText)(oFormSM.Items.Item("txtEmail2").Specific);

                    sCadenaCorreos = txtEmail1.Value.ToString();

                    if (string.IsNullOrEmpty(txtEmail2.Value.ToString()))
                    {

                    }
                    else
                    {
                        sCadenaCorreos = sCadenaCorreos + ";" + txtEmail2.Value.ToString();

                    }

                    DllFunciones.liberarObjetos(txtEmail1);
                    DllFunciones.liberarObjetos(txtEmail2);

                    #endregion
                }
                else if (iQuantityEmails == 3)
                {
                    #region Obtiene la cadena de correos

                    SAPbouiCOM.EditText txtEmail1 = (SAPbouiCOM.EditText)(oFormSM.Items.Item("txtEmail1").Specific);
                    SAPbouiCOM.EditText txtEmail2 = (SAPbouiCOM.EditText)(oFormSM.Items.Item("txtEmail2").Specific);
                    SAPbouiCOM.EditText txtEmail3 = (SAPbouiCOM.EditText)(oFormSM.Items.Item("txtEmail2").Specific);

                    sCadenaCorreos = txtEmail1.Value.ToString();

                    if (string.IsNullOrEmpty(txtEmail2.Value.ToString()))
                    {

                    }
                    else
                    {
                        sCadenaCorreos = sCadenaCorreos + ";" + txtEmail2.Value.ToString();

                    }

                    if (string.IsNullOrEmpty(txtEmail3.Value.ToString()))
                    {

                    }
                    else
                    {
                        sCadenaCorreos = sCadenaCorreos + ";" + txtEmail3.Value.ToString();

                    }

                    #endregion
                }
                else if (iQuantityEmails == 4)
                {
                    #region Obtiene la cadena de correos

                    SAPbouiCOM.EditText txtEmail1 = (SAPbouiCOM.EditText)(oFormSM.Items.Item("txtEmail1").Specific);
                    SAPbouiCOM.EditText txtEmail2 = (SAPbouiCOM.EditText)(oFormSM.Items.Item("txtEmail2").Specific);
                    SAPbouiCOM.EditText txtEmail3 = (SAPbouiCOM.EditText)(oFormSM.Items.Item("txtEmail3").Specific);
                    SAPbouiCOM.EditText txtEmail4 = (SAPbouiCOM.EditText)(oFormSM.Items.Item("txtEmail4").Specific);

                    sCadenaCorreos = txtEmail1.Value.ToString();

                    if (string.IsNullOrEmpty(txtEmail2.Value.ToString()))
                    {

                    }
                    else
                    {
                        sCadenaCorreos = sCadenaCorreos + ";" + txtEmail2.Value.ToString();

                    }

                    if (string.IsNullOrEmpty(txtEmail3.Value.ToString()))
                    {

                    }
                    else
                    {
                        sCadenaCorreos = sCadenaCorreos + ";" + txtEmail3.Value.ToString();

                    }

                    if (string.IsNullOrEmpty(txtEmail4.Value.ToString()))
                    {

                    }
                    else
                    {
                        sCadenaCorreos = sCadenaCorreos + ";" + txtEmail4.Value.ToString();

                    }

                    #endregion
                }
                else if (iQuantityEmails == 5)
                {
                    #region Obtiene la cadena de correos
                    SAPbouiCOM.EditText txtEmail1 = (SAPbouiCOM.EditText)(oFormSM.Items.Item("txtEmail1").Specific);
                    SAPbouiCOM.EditText txtEmail2 = (SAPbouiCOM.EditText)(oFormSM.Items.Item("txtEmail2").Specific);
                    SAPbouiCOM.EditText txtEmail3 = (SAPbouiCOM.EditText)(oFormSM.Items.Item("txtEmail3").Specific);
                    SAPbouiCOM.EditText txtEmail4 = (SAPbouiCOM.EditText)(oFormSM.Items.Item("txtEmail4").Specific);
                    SAPbouiCOM.EditText txtEmail5 = (SAPbouiCOM.EditText)(oFormSM.Items.Item("txtEmail5").Specific);

                    sCadenaCorreos = txtEmail1.Value.ToString();

                    if (string.IsNullOrEmpty(txtEmail2.Value.ToString()))
                    {

                    }
                    else
                    {
                        sCadenaCorreos = sCadenaCorreos + ";" + txtEmail2.Value.ToString();

                    }

                    if (string.IsNullOrEmpty(txtEmail3.Value.ToString()))
                    {

                    }
                    else
                    {
                        sCadenaCorreos = sCadenaCorreos + ";" + txtEmail3.Value.ToString();

                    }

                    if (string.IsNullOrEmpty(txtEmail4.Value.ToString()))
                    {

                    }
                    else
                    {
                        sCadenaCorreos = sCadenaCorreos + ";" + txtEmail4.Value.ToString();

                    }

                    if (string.IsNullOrEmpty(txtEmail5.Value.ToString()))
                    {

                    }
                    else
                    {
                        sCadenaCorreos = sCadenaCorreos + ";" + txtEmail5.Value.ToString();

                    }
                    #endregion
                }

                SendEmailResponse RespuestaEnvioCorreo = serviceClient.EnvioCorreo(Convert.ToString(oLlaveyPassword.Fields.Item("Llave").Value.ToString()), Convert.ToString(oLlaveyPassword.Fields.Item("Password").Value.ToString()), olblL1.Caption.ToString(), sCadenaCorreos, null);

                if (RespuestaEnvioCorreo.codigo == 200)
                {
                    DllFunciones.sendMessageBox(_sboapp, RespuestaEnvioCorreo.mensaje);
                    btnCancel.Item.Click();
                }
                else
                {
                    DllFunciones.sendMessageBox(_sboapp, RespuestaEnvioCorreo.mensaje);
                }

                #endregion
            }
            catch (Exception ex)
            {

                DllFunciones.sendMessageBox(sboapp, ex.ToString());
            }



        }

        public bool validacionEnviarCorreo(SAPbouiCOM.Application sboapp)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            #region obtiene formulario correo 

            SAPbouiCOM.Form oFormSM;
            oFormSM = sboapp.Forms.Item("BO_SM");

            #endregion

            SAPbouiCOM.EditText txtEmail1 = (SAPbouiCOM.EditText)(oFormSM.Items.Item("txtEmail1").Specific);

            if (string.IsNullOrEmpty(txtEmail1.Value.ToString()))
            {
                DllFunciones.sendMessageBox(sboapp, "Debe ingresar almenos un correo electronico para poder enviar");
                return false;
            }
            else
            {
                return true;
            }


        }

        public string VersionDll()
        {
            try
            {

                Assembly Assembly = Assembly.LoadFrom("BOeBilling.dll");
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
