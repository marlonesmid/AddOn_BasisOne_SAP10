using SAPbouiCOM;
using System;
using System.IO;
using System.Reflection;
using Funciones;
using SAPbobsCOM;
using System.Xml;
using System.Windows.Forms;


namespace Presupuesto
{
    public class Core
    {
        private SAPbouiCOM.Application sboapp;
        private SAPbobsCOM.Company oCompany;
       
        public Core(SAPbouiCOM.Application psboapp, SAPbobsCOM.Company _company)
        {
            this.sboapp = psboapp;
            this.oCompany = _company;
        }

        public void showForm(String menu)
        {
            Funciones.Comunes oFunc = new Funciones.Comunes();
            try
            {
                String frmResource1 = "";
                switch (menu)
                {
                    case "BO_ConfPresup":
                        //Llamar metodo crear formulario por SRF               
                        frmResource1 = $"{Assembly.GetExecutingAssembly().GetName().Name}.Formularios.ConfPresup.srf";
                        break;
                    case "BO_PerfilPresup":
                        //Llamar metodo crear formulario por SRF               
                        frmResource1 = $"{Assembly.GetExecutingAssembly().GetName().Name}.Formularios.PerfilPresup.srf";
                        break;
                    case "BO_PresupCuenta": //presupuesto por cuenta
                        //Llamar metodo crear formulario por SRF               
                        frmResource1 = $"{Assembly.GetExecutingAssembly().GetName().Name}.Formularios.PresupCuenta.srf";
                        menu = "BOPC";
                        break;
                    default:
                        break;
                }
                oFunc.crearFormPorXML(sboapp, getStream(frmResource1));
                var oForm = sboapp.Forms.Item(menu);
                // Mostrar formulario centrado
                oForm.Left = (sboapp.Desktop.Width - oForm.Width) / 2;
                oForm.Top = (sboapp.Desktop.Height - oForm.Height) / 4;
                oForm.Visible = true;
            }
            catch (Exception e)
            {
                oFunc.sendErrorMessage(sboapp, e);
            }
        }

        //metodo para generar el Stream que necesita el XML para crear el formulario
        private System.IO.Stream getStream(String frmResource)
        {
            Funciones.Comunes oFunc = new Funciones.Comunes();
            try
            {
                System.IO.Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(frmResource);
                return stream;
            }
            catch (Exception e)
            {
                oFunc.sendErrorMessage(sboapp, e);
                return null;
            }
        }
        
        public void creaTablasPresup(SAPbobsCOM.Company oCompany, SAPbouiCOM.Application sboapp)
        {
            Funciones.Comunes oFunc = new Funciones.Comunes();
            SAPbouiCOM.Form oForm = null;
            oForm = sboapp.Forms.Item("BO_ConfPresup");

            oForm = sboapp.Forms.ActiveForm;
            EditText oEditTxt = (EditText)oForm.Items.Item("BO_Txt1").Specific;
            StaticText oLbl = (StaticText)oForm.Items.Item("BO_Lbl1").Specific;
            try
            {
                sboapp.MetadataAutoRefresh = false;
                oLbl.Caption = "Creando Tablas";
                //Crear Tablas de Usuario Presupuesto por Cuenta
                oEditTxt.Value = "BO_PresupCuentaEnc";
                oFunc.crearTabla(oCompany, sboapp, "BO_PresupCuentaEnc", "BO Presupuesto Cuenta Enc", SAPbobsCOM.BoUTBTableType.bott_Document);
                oEditTxt.Value = "BO_PresupCuentaDet1";
                oFunc.crearTabla(oCompany, sboapp, "BO_PresupCuentaDet1", "BO Presupuesto Cuenta Det1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                oEditTxt.Value = "BO_PresupCuentaDet2";
                oFunc.crearTabla(oCompany, sboapp,"BO_PresupCuentaDet2", "BO Presupuesto Cuenta Det2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                oEditTxt.Value = "BO_PresupCuentaDet3";
                oFunc.crearTabla(oCompany, sboapp, "BO_PresupCuentaDet3", "BO Presupuesto Cuenta Det3", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);

                //Crear Campos de Usuario Presupuesto por Cuenta Encabezado
                int n = 19;
                oLbl.Caption = "Campos de usuario Enc...";
                oEditTxt.Value = $"1/{n}";
                oFunc.CreaCamposUsr(oCompany,  sboapp,BoFieldTypes.db_Alpha, BoFldSubTypes.st_None,25,"",BoYesNoEnum.tNO,null,"BO_PresupCuentaEnc","BO_Usuario","Usuario");
                oEditTxt.Value = $"2/{n}";
                oFunc.CreaCamposUsr(oCompany,  sboapp, BoFieldTypes.db_Date, BoFldSubTypes.st_None, 10, "", BoYesNoEnum.tYES, null, "BO_PresupCuentaEnc", "BO_FechaCrea", "Fecha Creación");
                oEditTxt.Value = $"3/{n}";
                oFunc.CreaCamposUsr(oCompany,  sboapp, BoFieldTypes.db_Numeric, BoFldSubTypes.st_None, 4, "", BoYesNoEnum.tYES, null, "BO_PresupCuentaEnc", "BO_Ano", "Año Vigencia");
                String[] oValidValues = { "Y", "Y", "N", "N"};
                oEditTxt.Value = $"4/{n}";
                oFunc.CreaCamposUsr(oCompany,  sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, "", BoYesNoEnum.tNO, oValidValues, "BO_PresupCuentaEnc", "BO_ProyectoSN", "ProyectoSN");
                oEditTxt.Value = $"5/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 20, "", BoYesNoEnum.tNO, null, "BO_PresupCuentaEnc", "BO_Proyecto", "Proyecto");
                oEditTxt.Value = $"6/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, "", BoYesNoEnum.tNO, oValidValues, "BO_PresupCuentaEnc", "BO_Dimension1SN", "Dimension1SN");
                oEditTxt.Value = $"7/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 8, "", BoYesNoEnum.tNO, null, "BO_PresupCuentaEnc", "BO_Dimension1", "Dimension1");
                oEditTxt.Value = $"8/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, "", BoYesNoEnum.tNO, oValidValues, "BO_PresupCuentaEnc", "BO_Dimension2SN", "Dimension1SN");
                oEditTxt.Value = $"9/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 8, "", BoYesNoEnum.tNO, null, "BO_PresupCuentaEnc", "BO_Dimension2", "Dimension2");
                oEditTxt.Value = $"10/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, "", BoYesNoEnum.tNO, oValidValues, "BO_PresupCuentaEnc", "BO_Dimension3SN", "Dimension3SN");
                oEditTxt.Value = $"11/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 8, "", BoYesNoEnum.tNO, null, "BO_PresupCuentaEnc", "BO_Dimension3", "Dimension3");
                oEditTxt.Value = $"12/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, "", BoYesNoEnum.tNO, oValidValues, "BO_PresupCuentaEnc", "BO_Dimension4SN", "Dimension4SN");
                oEditTxt.Value = $"13/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 8, "", BoYesNoEnum.tNO, null, "BO_PresupCuentaEnc", "BO_Dimension4", "Dimension4");
                oEditTxt.Value = $"14/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, "", BoYesNoEnum.tNO, oValidValues, "BO_PresupCuentaEnc", "BO_Dimension5SN", "Dimension5SN");
                oEditTxt.Value = $"15/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 8, "", BoYesNoEnum.tNO, null, "BO_PresupCuentaEnc", "BO_Dimension5", "Dimension5");
                oEditTxt.Value = $"16/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Memo, BoFldSubTypes.st_None, 255, "", BoYesNoEnum.tNO, null, "BO_PresupCuentaEnc", "BO_Nombre", "Nombre");
                oEditTxt.Value = $"17/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, "Y", BoYesNoEnum.tYES, oValidValues, "BO_PresupCuentaEnc", "BO_Activo", "Activo");
                String[] oValidValuesStatus = { "B", "Borrador", "L", "Liberado", "A", "Aprobado" };
                oEditTxt.Value = $"18/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 15, "B", BoYesNoEnum.tYES, oValidValuesStatus, "BO_PresupCuentaEnc", "BO_Status", "Status");
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 5, "", BoYesNoEnum.tNO, null, "BO_PresupCuentaEnc", "BO_UsrId", "Id");
                oEditTxt.Value = $"19/{n}";

                //Crear campos Presupuesto Cuenta Detalle1
                oLbl.Caption = "Campos de usuario Det1...";
                oEditTxt.Value = $"1/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 15, "", BoYesNoEnum.tNO, null, "BO_PresupCuentaDet1", "BO_Cuenta", "Cuenta");
                String[] oValid1 = {"D", "Débito", "C", "Crédito" };
                oEditTxt.Value = $"2/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, "D", BoYesNoEnum.tNO, oValid1
                    , "BO_PresupCuentaDet1", "BO_Naturaleza", "Naturaleza");
                oEditTxt.Value = $"3/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Float, BoFldSubTypes.st_Price, 11, "0.0", BoYesNoEnum.tNO, null
                    , "BO_PresupCuentaDet1", "BO_Enero", "Enero");
                oEditTxt.Value = $"4/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Float, BoFldSubTypes.st_Price, 11, "0.0", BoYesNoEnum.tNO, null
                    , "BO_PresupCuentaDet1", "BO_Febrero", "Febrero");
                oEditTxt.Value = $"5/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Float, BoFldSubTypes.st_Price, 11, "0.0", BoYesNoEnum.tNO, null
                    , "BO_PresupCuentaDet1", "BO_Marzo", "Marzo");
                oEditTxt.Value = $"6/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Float, BoFldSubTypes.st_Price, 11, "0.0", BoYesNoEnum.tNO, null
                    , "BO_PresupCuentaDet1", "BO_Abril", "Abril");
                oEditTxt.Value = $"7/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Float, BoFldSubTypes.st_Price, 11, "0.0", BoYesNoEnum.tNO, null
                    , "BO_PresupCuentaDet1", "BO_Mayo", "Mayo");
                oEditTxt.Value = $"8/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Float, BoFldSubTypes.st_Price, 11, "0.0", BoYesNoEnum.tNO, null
                    , "BO_PresupCuentaDet1", "BO_Junio", "Junio");
                oEditTxt.Value = $"9/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Float, BoFldSubTypes.st_Price, 11, "0.0", BoYesNoEnum.tNO, null
                    , "BO_PresupCuentaDet1", "BO_Julio", "Julio");
                oEditTxt.Value = $"10/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Float, BoFldSubTypes.st_Price, 11, "0.0", BoYesNoEnum.tNO, null
                    , "BO_PresupCuentaDet1", "BO_Agosto", "Agosto");
                oEditTxt.Value = $"11/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Float, BoFldSubTypes.st_Price, 11, "0.0", BoYesNoEnum.tNO, null
                    , "BO_PresupCuentaDet1", "BO_Septiembre", "Septiembre");
                oEditTxt.Value = $"12/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Float, BoFldSubTypes.st_Price, 11, "0.0", BoYesNoEnum.tNO, null
                    , "BO_PresupCuentaDet1", "BO_Octubre", "Octubre");
                oEditTxt.Value = $"13/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Float, BoFldSubTypes.st_Price, 11, "0.0", BoYesNoEnum.tNO, null
                    , "BO_PresupCuentaDet1", "BO_Noviembre", "Noviembre");
                oEditTxt.Value = $"14/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Float, BoFldSubTypes.st_Price, 11, "0.0", BoYesNoEnum.tNO, null
                    , "BO_PresupCuentaDet1", "BO_Diciembre", "Diciembre");
                oEditTxt.Value = $"15/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Float, BoFldSubTypes.st_Price, 11, "0.0", BoYesNoEnum.tNO, null
                    , "BO_PresupCuentaDet1", "BO_Total", "Total");
                oEditTxt.Value = $"16/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Memo, BoFldSubTypes.st_None, 255, "", BoYesNoEnum.tNO, null, "BO_PresupCuentaDet1", "BO_Comentario", "Comentario");
                oEditTxt.Value = $"17/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, "", BoYesNoEnum.tNO, oValidValues, "BO_PresupCuentaDet1", "BO_Activo", "Activo");

                //Crear campos Presupuesto Cuenta Detalle2
                oLbl.Caption = "Campos de usuario Det2...";
                oEditTxt.Value = $"1/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 15, "", BoYesNoEnum.tNO, null, "BO_PresupCuentaDet2", "BO_Cuenta", "Cuenta");
                oEditTxt.Value = $"2/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, "D", BoYesNoEnum.tNO, oValid1
                    , "BO_PresupCuentaDet2", "BO_Naturaleza", "Naturaleza");
                oEditTxt.Value = $"3/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Float, BoFldSubTypes.st_Price, 11, "0.0", BoYesNoEnum.tNO, null
                    , "BO_PresupCuentaDet2", "BO_Enero", "Enero");
                oEditTxt.Value = $"4/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Float, BoFldSubTypes.st_Price, 11, "0.0", BoYesNoEnum.tNO, null
                    , "BO_PresupCuentaDet2", "BO_Febrero", "Febrero");
                oEditTxt.Value = $"5/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Float, BoFldSubTypes.st_Price, 11, "0.0", BoYesNoEnum.tNO, null
                    , "BO_PresupCuentaDet2", "BO_Marzo", "Marzo");
                oEditTxt.Value = $"6/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Float, BoFldSubTypes.st_Price, 11, "0.0", BoYesNoEnum.tNO, null
                    , "BO_PresupCuentaDet2", "BO_Abril", "Abril");
                oEditTxt.Value = $"7/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Float, BoFldSubTypes.st_Price, 11, "0.0", BoYesNoEnum.tNO, null
                    , "BO_PresupCuentaDet2", "BO_Mayo", "Mayo");
                oEditTxt.Value = $"8/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Float, BoFldSubTypes.st_Price, 11, "0.0", BoYesNoEnum.tNO, null
                    , "BO_PresupCuentaDet2", "BO_Junio", "Junio");
                oEditTxt.Value = $"9/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Float, BoFldSubTypes.st_Price, 11, "0.0", BoYesNoEnum.tNO, null
                    , "BO_PresupCuentaDet2", "BO_Julio", "Julio");
                oEditTxt.Value = $"10/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Float, BoFldSubTypes.st_Price, 11, "0.0", BoYesNoEnum.tNO, null
                    , "BO_PresupCuentaDet2", "BO_Agosto", "Agosto");
                oEditTxt.Value = $"11/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Float, BoFldSubTypes.st_Price, 11, "0.0", BoYesNoEnum.tNO, null
                    , "BO_PresupCuentaDet2", "BO_Septiembre", "Septiembre");
                oEditTxt.Value = $"12/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Float, BoFldSubTypes.st_Price, 11, "0.0", BoYesNoEnum.tNO, null
                    , "BO_PresupCuentaDet2", "BO_Octubre", "Octubre");
                oEditTxt.Value = $"13/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Float, BoFldSubTypes.st_Price, 11, "0.0", BoYesNoEnum.tNO, null
                    , "BO_PresupCuentaDet2", "BO_Noviembre", "Noviembre");
                oEditTxt.Value = $"14/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Float, BoFldSubTypes.st_Price, 11, "0.0", BoYesNoEnum.tNO, null
                    , "BO_PresupCuentaDet2", "BO_Diciembre", "Diciembre");
                oEditTxt.Value = $"15/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Float, BoFldSubTypes.st_Price, 11, "0.0", BoYesNoEnum.tNO, null
                    , "BO_PresupCuentaDet2", "BO_Total", "Total");
                oEditTxt.Value = $"16/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Memo, BoFldSubTypes.st_None, 0, "", BoYesNoEnum.tNO, null, "BO_PresupCuentaDet2", "BO_Comentario", "Comentario");
                oEditTxt.Value = $"17/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, "", BoYesNoEnum.tNO, oValidValues, "BO_PresupCuentaDet2", "BO_Activo", "Activo");

                //Crear campos Presupuesto Cuenta Detalle3
                oLbl.Caption = "Campos de usuario Det3...";
                n = 3;
                oEditTxt.Value = $"1/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Memo, BoFldSubTypes.st_Link, 0, "", BoYesNoEnum.tNO, null, "BO_PresupCuentaDet3", "BO_Anexo", "Anexo");
                oEditTxt.Value = $"2/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Memo, BoFldSubTypes.st_None, 0, "", BoYesNoEnum.tNO, null, "BO_PresupCuentaDet3", "BO_Comentario", "Comentario");
                oEditTxt.Value = $"3/{n}";
                oFunc.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Date, BoFldSubTypes.st_None, 10, "", BoYesNoEnum.tNO, null, "BO_PresupCuentaDet3", "BO_Fecha", "Fecha");

                //Crear UDO
                oLbl.Caption = "Creando UDO PresupCuenta";
                oEditTxt.Value = "UDO PresupCuenta";
                string[] T1 = { "BO_PresupCuentaEnc", "BO_PresupCuentaDet1", "BO_PresupCuentaDet2", "BO_PresupCuentaDet3" };
                string[] F1 = { "DocEntry", "U_BO_Nombre", "U_BO_Usuario", "U_BO_Ano", "U_BO_Proyecto", "U_BO_Dimension1", "U_BO_Dimension2", "U_BO_Dimension3", "U_BO_Dimension4", "U_BO_Dimension5", "U_BO_Activo", "U_BO_Status" };
                oFunc.CrearUDO(oCompany,sboapp, "BOPC", "Presupuesto por Cuenta",BoUDOObjType.boud_Document, T1,BoYesNoEnum.tNO,BoYesNoEnum.tYES,F1,BoYesNoEnum.tNO,BoYesNoEnum.tNO,BoYesNoEnum.tNO,BoYesNoEnum.tNO,BoYesNoEnum.tNO,BoYesNoEnum.tNO,BoYesNoEnum.tNO,BoYesNoEnum.tNO,0,1,BoYesNoEnum.tYES,"BO_PresCuLog");

                oEditTxt.Value = "Finalizando proceso...";
                sboapp.MetadataAutoRefresh = true;
                //System.Threading.Thread.Sleep(2000);

            }
            catch (Exception e)
            {
                oFunc.sendErrorMessage(sboapp, e);
            }
            finally
            {
                // oForm.Close();
                oFunc.sendMessageBox(sboapp, "Proceso finalizado con éxito", 2);
            }           
        }
     
        public void LlenarChkForms(string oForm)
        {
            if (oForm == "BOPC")
            {
                SAPbouiCOM.Form oFormBOPC = null;
                oFormBOPC = sboapp.Forms.Item("BOPC");
                //oForm = sboapp.Forms.ActiveForm;
                SAPbouiCOM.CheckBox chk01 = (SAPbouiCOM.CheckBox)oFormBOPC.Items.Item("BO_ProySN").Specific;
                SAPbouiCOM.CheckBox chk02 = (SAPbouiCOM.CheckBox)oFormBOPC.Items.Item("BO_Dim1SN").Specific;
                SAPbouiCOM.CheckBox chk03 = (SAPbouiCOM.CheckBox)oFormBOPC.Items.Item("BO_Dim2SN").Specific;
                SAPbouiCOM.CheckBox chk04 = (SAPbouiCOM.CheckBox)oFormBOPC.Items.Item("BO_Dim3SN").Specific;
                SAPbouiCOM.CheckBox chk05 = (SAPbouiCOM.CheckBox)oFormBOPC.Items.Item("BO_Dim4SN").Specific;
                SAPbouiCOM.CheckBox chk06 = (SAPbouiCOM.CheckBox)oFormBOPC.Items.Item("BO_Dim5SN").Specific;
                EditText txt01 = (EditText)oFormBOPC.Items.Item("BO_Proy").Specific;
                EditText txt02 = (EditText)oFormBOPC.Items.Item("BO_Dim1").Specific;
                EditText txt03 = (EditText)oFormBOPC.Items.Item("BO_Dim2").Specific;
                EditText txt04 = (EditText)oFormBOPC.Items.Item("BO_Dim3").Specific;
                EditText txt05 = (EditText)oFormBOPC.Items.Item("BO_Dim4").Specific;
                EditText txt06 = (EditText)oFormBOPC.Items.Item("BO_Dim5").Specific;
                //ComboBox cbx01 = (ComboBox)oFormBOPC.Items.Item("BO_Activo").Specific;
                //ComboBox cbx02 = (ComboBox)oFormBOPC.Items.Item("BO_Status").Specific;

                //cbx01.Select("Y");
                //cbx02.Select("B");
                chk01.Checked = false;
                chk02.Checked = false;
                chk03.Checked = false;
                chk04.Checked = false;
                chk05.Checked = false;
                chk06.Checked = false;
                
                txt01.Item.Enabled = false;
                txt02.Item.Enabled = false;
                txt03.Item.Enabled = false;
                txt04.Item.Enabled = false;
                txt05.Item.Enabled = false;
                txt06.Item.Enabled = false;

            }

        }

        public void LimpiarChkForm(string oForm, IItemEvent campo)
        {
            if (oForm == "BOPC")
            {
                SAPbouiCOM.Form oFormBOPC = null;
                oFormBOPC = sboapp.Forms.Item("BOPC");
                SAPbouiCOM.CheckBox chk01 = (SAPbouiCOM.CheckBox)oFormBOPC.Items.Item("BO_ProySN").Specific;
                SAPbouiCOM.CheckBox chk02 = (SAPbouiCOM.CheckBox)oFormBOPC.Items.Item("BO_Dim1SN").Specific;
                SAPbouiCOM.CheckBox chk03 = (SAPbouiCOM.CheckBox)oFormBOPC.Items.Item("BO_Dim2SN").Specific;
                SAPbouiCOM.CheckBox chk04 = (SAPbouiCOM.CheckBox)oFormBOPC.Items.Item("BO_Dim3SN").Specific;
                SAPbouiCOM.CheckBox chk05 = (SAPbouiCOM.CheckBox)oFormBOPC.Items.Item("BO_Dim4SN").Specific;
                SAPbouiCOM.CheckBox chk06 = (SAPbouiCOM.CheckBox)oFormBOPC.Items.Item("BO_Dim5SN").Specific;

                EditText txtID = (EditText)oFormBOPC.Items.Item("BO_User").Specific;
                EditText txt01 = (EditText)oFormBOPC.Items.Item("BO_Proy").Specific;
                EditText txt02 = (EditText)oFormBOPC.Items.Item("BO_Dim1").Specific;
                EditText txt03 = (EditText)oFormBOPC.Items.Item("BO_Dim2").Specific;
                EditText txt04 = (EditText)oFormBOPC.Items.Item("BO_Dim3").Specific;
                EditText txt05 = (EditText)oFormBOPC.Items.Item("BO_Dim4").Specific;
                EditText txt06 = (EditText)oFormBOPC.Items.Item("BO_Dim5").Specific;
                

                

                txtID.Item.Click();
                switch (campo.ItemUID)
                {
                    case "BO_ProySN":
                        if (chk01.Checked == true)
                        {
                            txt01.Item.Enabled = true;
                        }
                        else
                        {
                            txt01.Item.Enabled = false;
                        }
                        break;
                    case "BO_Dim1SN":
                        if (chk02.Checked == true)
                        {
                            txt02.Item.Enabled = true;
                        }
                        else
                        {
                            txt02.Item.Enabled = false;
                        }
                        break;
                    case "BO_Dim2SN":
                        if (chk03.Checked == true)
                        {
                            txt03.Item.Enabled = true;
                        }
                        else
                        {
                            txt03.Item.Enabled = false;
                        }
                        break;
                    case "BO_Dim3SN":
                        if (chk04.Checked == true)
                        {
                            txt04.Item.Enabled = true;
                        }
                        else
                        {
                            txt04.Item.Enabled = false;
                        }
                        break;
                    case "BO_Dim4SN":
                        if (chk05.Checked == true)
                        {
                            txt05.Item.Enabled = true;
                        }
                        else
                        {
                            txt05.Item.Enabled = false;
                        }
                        break;
                    case "BO_Dim5SN":
                        if (chk06.Checked == true)
                        {
                            txt06.Item.Enabled = true;
                        }
                        else
                        {
                            txt06.Item.Enabled = false;
                        }
                        break;
                    default:
                        break;
                }
            }
        }

        public Boolean presupEventos(SAPbouiCOM.BoEventTypes evento, ItemEvent pVal)
        {
            Funciones.Comunes oFunc = new Funciones.Comunes();
            switch (evento)
            {               
                case BoEventTypes.et_CLICK:
                   
                    if (pVal.FormUID == "BOPC" && pVal.ItemUID == "1" && pVal.Before_Action == true && pVal.Action_Success == false)
                    {
                        if (pVal.FormMode == 3)
                        {
                            SAPbouiCOM.Form oFormBOPC = null;
                            oFormBOPC = sboapp.Forms.Item("BOPC");
                            EditText txt01 = (EditText)oFormBOPC.Items.Item("BO_User").Specific;
                            EditText txt02 = (EditText)oFormBOPC.Items.Item("BO_Nombre").Specific;
                            EditText txt03 = (EditText)oFormBOPC.Items.Item("BO_Fech").Specific;
                            EditText txt04 = (EditText)oFormBOPC.Items.Item("BO_Ano").Specific;
                            SAPbouiCOM.ComboBox cbx01 = (SAPbouiCOM.ComboBox)oFormBOPC.Items.Item("BO_Activo").Specific;
                            SAPbouiCOM.ComboBox cbx02 = (SAPbouiCOM.ComboBox)oFormBOPC.Items.Item("BO_Status").Specific;

                            if (String.IsNullOrEmpty(txt01.Value))
                            {
                                oFunc.sendMessageBox(sboapp, "El campo Usuario es obligatorio", 1);
                                txt01.Item.Click();
                                return false;
                            }
                            else if (String.IsNullOrEmpty(txt02.Value))
                            {
                                oFunc.sendMessageBox(sboapp, "El campo Nombre del Presupuesto es obligatorio", 1);
                                txt02.Item.Click();
                                return false;
                            }
                            else if (String.IsNullOrEmpty(txt03.Value))
                            {
                                oFunc.sendMessageBox(sboapp, "El campo Fecha es obligatorio", 1);
                                txt03.Item.Click();
                                return false;
                            }
                            else if (String.IsNullOrEmpty(txt04.Value))
                            {
                                oFunc.sendMessageBox(sboapp, "El campo Año es obligatorio", 1);
                                txt04.Item.Click();
                                return false;
                            }
                            else if (txt04.Value.Length >4 || txt04.Value.Length < 4)
                            {
                                oFunc.sendMessageBox(sboapp, "Campo Año debe contener 4 dígitos", 1);
                                txt04.Item.Click();
                                return false;
                            }
                            else if (txt04.Value.Length == 4 && (txt04.Value.Substring(0, 2) != "20"))
                            {       
                                oFunc.sendMessageBox(sboapp, "El Año debe debe se mayor al 2018", 1);
                                txt04.Item.Click();
                                return false;                               
                            }
                            else if (String.IsNullOrEmpty(cbx01.Value))
                            {
                                oFunc.sendMessageBox(sboapp, "El campo Activo/Inactivo es obligatorio", 1);
                                cbx01.Item.Click();
                                return false;
                            }
                            else if (String.IsNullOrEmpty(cbx02.Value))
                            {
                                oFunc.sendMessageBox(sboapp, "El campo Estado es obligatorio", 1);
                                cbx02.Item.Click();
                                return false;
                            }
                            else
                            {
                                return true;
                            }

                        }
                    }
                    break;
               
                case BoEventTypes.et_FORM_ACTIVATE:
                    if (pVal.FormUID == "BOPC")
                    {
                        if (pVal.FormMode == 3) //from ADD_MODE
                        {
                            string usrID;
                            string usrName;
                            SAPbouiCOM.Form oForm = null;
                            oForm = sboapp.Forms.Item("BOPC");
                            EditText et01 = (EditText)oForm.Items.Item("BO_User").Specific;
                            EditText et02 = (EditText)oForm.Items.Item("BO_UserID").Specific;
                            
                            usrID = oFunc.GetUsrID(oCompany, sboapp, et01.Value);
                            usrName = oFunc.GetUsrName(oCompany, sboapp, et01.Value);
                            et02.Value = usrID;
                            et01.Item.Click();
                            return true;
                        }
                    }
                    break;
           
                case BoEventTypes.et_FORM_VISIBLE:
                    if (pVal.FormUID == "BOPC")
                    {
                        SAPbouiCOM.Form oForm = null;
                        oForm = sboapp.Forms.Item("BOPC");
                        return true;
                    }
                    
                    break;
         
                default:
                    return true;
                    break;

            }
            return true;
        }

        public string VersionDll()
        {
            try
            {

                Assembly Assembly = Assembly.LoadFrom("BOPresupuesto.dll");
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
