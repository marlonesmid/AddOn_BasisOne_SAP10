using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbobsCOM;
using SAPbouiCOM;
using Funciones;
using System.IO;
using System.Reflection;


namespace BOProduccion
{
    public class Production
    {

        public void CreateUDTandUDFProduction(SAPbouiCOM.Application sboapp, SAPbobsCOM.Company oCompany)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            #region Creacion de tablas

            DllFunciones.ProgressBar(oCompany, sboapp, 16, 1, "Creando Tabla - Parametros Produccion Avanzada, por favor espere...");
            DllFunciones.crearTabla(oCompany, sboapp, "BOPRODP", "BO-Param. Produc. Avan.", SAPbobsCOM.BoUTBTableType.bott_NoObject);

            DllFunciones.ProgressBar(oCompany, sboapp, 16, 1, "Creando Tabla - BORTDC - BO Resgistro de tiempo detallado , por favor espere...");
            DllFunciones.crearTabla(oCompany, sboapp, "BORTDC", "BO-Reg.Tiem.Deta.Ca", SAPbobsCOM.BoUTBTableType.bott_Document);

            DllFunciones.ProgressBar(oCompany, sboapp, 16, 1, "Creando Tabla - BORTDD - BO Resgistro de tiempo detallado , por favor espere...");
            DllFunciones.crearTabla(oCompany, sboapp, "BORTDD", "BO-Reg.Tiem.Deta.De", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);

            #endregion

            #region Creacion de Campos

            DllFunciones.ProgressBar(oCompany, sboapp, 16, 1, "Creando Campo - OWOR - Tipo de Orden ...");
            string[] ValidValuesFields1 = { "T", "Prodcuto Terminado", "S", "Producto Semielaborado" };
            DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, "", BoYesNoEnum.tNO, ValidValuesFields1, "OWOR", "BO_TO", "Tipo Orden");

            DllFunciones.ProgressBar(oCompany, sboapp, 16, 1, "Creando Campo - OWOR - OP Principal.. ");
            DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 30, "", BoYesNoEnum.tNO, null, "OWOR", "BO_OPP", "OP Principal");

            DllFunciones.ProgressBar(oCompany, sboapp, 16, 1, "Creando Campo - OWOR - Posicion Articulo.. ");
            DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 30, "", BoYesNoEnum.tNO, null, "OWOR", "BO_PosId", "Posicion OP");

            //DllFunciones.ProgressBar(oCompany, sboapp, 16, 1, "Creando Campo - WOR1 - Registro detallado de tiempo.. ");
            //DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 10, "", BoYesNoEnum.tNO, null, "WOR1", "BO_RTD", "Res. Tiem. Deta");

            DllFunciones.ProgressBar(oCompany, sboapp, 16, 1, "Creando Campo - BOPRODP - Serie numeracion Produc. Terminado... ");
            DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 30, "", BoYesNoEnum.tNO, null, "@BOPRODP", "BO_SNPT", "Ser.Num.PP");

            DllFunciones.ProgressBar(oCompany, sboapp, 16, 1, "Creando Campo - BOPRODP - Serie numeracion Produc. Semielaborado... ");
            DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 30, "", BoYesNoEnum.tNO, null, "@BOPRODP", "BO_SNPS", "Ser.Num.PS");

            DllFunciones.ProgressBar(oCompany, sboapp, 16, 1, "Creando Campo - BOPRODP - Ruta Imagenes ... ");
            DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "@BOPRODP", "BO_RIMG", "Ruta Imagenes");

            DllFunciones.ProgressBar(oCompany, sboapp, 16, 1, "Creando Campo - BORTDD - Persona  ... ");
            DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 100, "", BoYesNoEnum.tNO, null, "@BORTDD", "BO_P", "Persona");

            DllFunciones.ProgressBar(oCompany, sboapp, 16, 1, "Creando Campo - BORTDD - Nombre Persona  ... ");
            DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "@BORTDD", "BO_NP", "Nombre Persona");

            DllFunciones.ProgressBar(oCompany, sboapp, 16, 1, "Creando Campo - BORTDD - Fecha Registro... ");
            DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Date, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "@BORTDD", "BO_FR", "Fecha Registro");

            DllFunciones.ProgressBar(oCompany, sboapp, 16, 1, "Creando Campo - BORTDD - Hora desde ... ");
            DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Date, BoFldSubTypes.st_Time, 254, "", BoYesNoEnum.tNO, null, "@BORTDD", "BO_TI", "Hora desde");

            DllFunciones.ProgressBar(oCompany, sboapp, 16, 1, "Creando Campo - BORTDD - Hora Hasta ... ");
            DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Date, BoFldSubTypes.st_Time, 254, "", BoYesNoEnum.tNO, null, "@BORTDD", "BO_TF", "Hora hasta");


            #endregion

            #region Creacion de UDOS 

            DllFunciones.ProgressBar(oCompany, sboapp, 16, 1, "Creando UDO - Registro de tiempo detallado..");
            string[] TablaseBilling = { "BORTDC", "BORTDD" };
            DllFunciones.CrearUDO(oCompany, sboapp, "BORTD", "BO Registro Tiempos", BoUDOObjType.boud_Document, TablaseBilling, BoYesNoEnum.tNO, BoYesNoEnum.tYES, null, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, 0, 1, BoYesNoEnum.tYES, "BO_BORTD_Log");

            #endregion

            #region Creacion de procedimientos almacenados

            DllFunciones.ProgressBar(oCompany, sboapp, 16, 1, "Creando Procedimientos almacenados , por favor espere...");
            SAPbobsCOM.Recordset oProcedures = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            #region Consulta si el procedure Existe

            string sProcedure_Eliminar;
            string sProcedure_Crear;

            sProcedure_Eliminar = DllFunciones.GetStringXMLDocument(oCompany, "BOProduction", "Production", "DeleteProcedure");
            sProcedure_Eliminar = sProcedure_Eliminar.Replace("%sNameProcedure%", "BO_OrdenesProduccion");

            oProcedures.DoQuery(sProcedure_Eliminar);

            #endregion

            #region Crea el procedure

            sProcedure_Crear = DllFunciones.GetStringXMLDocument(oCompany, "BOProduction", "Production", "ProcedureWorkOrder");

            string sPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\BOProduction\\Images\\";

            sProcedure_Crear = sProcedure_Crear.Replace("%sPath%", sPath);

            oProcedures.DoQuery(sProcedure_Crear);

            DllFunciones.liberarObjetos(oProcedures);
            sProcedure_Crear = null;
            sProcedure_Eliminar = null;

            #endregion

            #endregion

        }

        public Boolean Validate_WorkOrder(SAPbouiCOM.Application sboapp, SAPbobsCOM.Company oCompany, SAPbouiCOM.Form oFormWorkOrder)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            try
            {
                int iContinuar = DllFunciones.sendMessageBoxY_N(sboapp, "Se creara la ruta de produccion de la orden de fabricacion, ¿ Desea Crear ?");

                if (iContinuar == 1)
                {
                    return true;
                }
                else
                {
                    return false;
                }

            }
            catch (Exception)
            {

                throw;
                return false;
            }

        }

        public Boolean Create_Order_Prodcution(SAPbouiCOM.Application sboapp, SAPbobsCOM.Company oCompany, string _sMotor, string sUserSignature)
        {
            #region Variables globales

            Boolean Flag;

            #endregion

            #region Intanciacion de Dll's

            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            #endregion

            try
            {
                #region Variables y objetos

                string sGetWorkOrder;

                SAPbobsCOM.Recordset oGetWorkOrder = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

                #endregion

                #region Valida si existe la orden de produccion de producto terminado para crear los productos semielaborados

                sGetWorkOrder = DllFunciones.GetStringXMLDocument(oCompany, "BOProduction", "Production", "GetWorkOrder");
                sGetWorkOrder = sGetWorkOrder.Replace("%UserSign%", sUserSignature);

                oGetWorkOrder.DoQuery(sGetWorkOrder);

                #endregion

                if (oGetWorkOrder.RecordCount > 0)
                {
                    #region Variable y Objetos

                    string sGetWorkOrderLine;
                    string sSearchBOMOriginal;

                    SAPbobsCOM.Recordset oGetWorkOrderLine = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

                    #endregion

                    #region Buscamos la Lineas de las ordenes de produccion

                    sGetWorkOrderLine = DllFunciones.GetStringXMLDocument(oCompany, "BOProduction", "Production", "GetWorkOrderLine");
                    sGetWorkOrderLine = sGetWorkOrderLine.Replace("%DocEntry%", Convert.ToString(oGetWorkOrder.Fields.Item("DocEntry").Value.ToString()));

                    oGetWorkOrderLine.DoQuery(sGetWorkOrderLine);

                    #endregion

                    #region Busca el Query de la lista de materiales 

                    sSearchBOMOriginal = DllFunciones.GetStringXMLDocument(oCompany, "BOProduction", "Production", "SearchBOM");

                    #endregion

                    if (oGetWorkOrderLine.RecordCount > 0)
                    {

                        DllFunciones.ProgressBar(oCompany, sboapp, oGetWorkOrderLine.RecordCount + 1, 1, "Creando ruta de producion, por favor espere...");

                        #region Variables y objetos

                        string oSearchBOMCopia;

                        SAPbobsCOM.Recordset oConsultaBOM = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        SAPbobsCOM.ProductionOrders oWorkOrder = (SAPbobsCOM.ProductionOrders)oCompany.GetBusinessObject(BoObjectTypes.oProductionOrders);

                        #endregion

                        #region Crea las ruta de produccion

                        oGetWorkOrderLine.MoveFirst();

                        for (int i = 0; oGetWorkOrderLine.RecordCount - 1 >= i; i++)
                        {
                            #region Valida si existe la lista de materiales 

                            oSearchBOMCopia = sSearchBOMOriginal.Replace("%Item%", Convert.ToString(oGetWorkOrderLine.Fields.Item("ItemCode").Value.ToString()));

                            oConsultaBOM.DoQuery(oSearchBOMCopia);

                            #endregion

                            if (oConsultaBOM.RecordCount > 0)
                            {
                                #region Crea la orden de produccion semielaborada

                                oWorkOrder.ItemNo = Convert.ToString(oGetWorkOrderLine.Fields.Item("ItemCode").Value.ToString());
                                oWorkOrder.Series = Convert.ToInt32(oGetWorkOrderLine.Fields.Item("SNPS").Value.ToString());
                                oWorkOrder.StartDate = DateTime.Now;
                                oWorkOrder.ProductionOrderType = BoProductionOrderTypeEnum.bopotStandard;
                                oWorkOrder.PlannedQuantity = Convert.ToInt16(oGetWorkOrderLine.Fields.Item("PlannedQty").Value.ToString());
                                oWorkOrder.Warehouse = Convert.ToString(oGetWorkOrderLine.Fields.Item("wareHouse").Value.ToString());
                                oWorkOrder.UserFields.Fields.Item("U_BO_TO").Value = "S";
                                oWorkOrder.UserFields.Fields.Item("U_BO_OPP").Value = Convert.ToString(oGetWorkOrder.Fields.Item("DocNum").Value.ToString());
                                oWorkOrder.UserFields.Fields.Item("U_BO_PosId").Value = Convert.ToString(oGetWorkOrderLine.Fields.Item("VisOrder").Value.ToString());
                                int Rsd = oWorkOrder.Add();

                                if (Rsd == 0)
                                {
                                    DllFunciones.ProgressBar(oCompany, sboapp, oGetWorkOrderLine.RecordCount + 1, 1, "Creando ruta de producion, por favor espere...");
                                }
                                else
                                {
                                    DllFunciones.sendMessageBox(sboapp, oCompany.GetLastErrorDescription());
                                }

                                #endregion
                            }

                            oGetWorkOrderLine.MoveNext();
                        }

                        #endregion

                    }

                }
                DllFunciones.StatusBar(sboapp, BoStatusBarMessageType.smt_Success, "Posiciones creadas correctamente");
                return Flag = true;
            }
            catch (Exception e)
            {

                DllFunciones.sendErrorMessage(sboapp, e);
                Flag = false;
            }
            return Flag;

        }

        public void ChangueFormParProduction(SAPbouiCOM.Application sboapp, SAPbobsCOM.Company oCompany, SAPbouiCOM.Form oFormParProduction)
        {
            Funciones.Comunes DLLFunciones = new Funciones.Comunes();

            try
            {
                #region Variables y objetos

                SAPbouiCOM.ComboBox _cboSNPT = (SAPbouiCOM.ComboBox)oFormParProduction.Items.Item("txtSNPT").Specific;
                SAPbouiCOM.ComboBox _cboSNPS = (SAPbouiCOM.ComboBox)oFormParProduction.Items.Item("txtSNPS").Specific;

                string sNumberSeriesActive = null;
                string sNumberSeriesSAP = null;

                SAPbouiCOM.PictureBox oLogoBO = (SAPbouiCOM.PictureBox)oFormParProduction.Items.Item("imgLogoBO").Specific;
                SAPbobsCOM.Recordset oValidValuesSNActive = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset oValidValuesSNSAP = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                #endregion

                #region Centra en pantalla formulario

                oFormParProduction.Left = (sboapp.Desktop.Width - oFormParProduction.Width) / 2;
                oFormParProduction.Top = (sboapp.Desktop.Height - oFormParProduction.Height) / 4;

                #endregion

                #region Asignacion Logo BO

                oLogoBO.Picture = (Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Core\\Imagenes\\BO.jpg");

                #endregion

                #region Busqueda de series de numeracion asignada

                sNumberSeriesActive = DLLFunciones.GetStringXMLDocument(oCompany, "BOProduction", "Production", "GetNumberSerieActive");
                sNumberSeriesSAP = DLLFunciones.GetStringXMLDocument(oCompany, "BOProduction", "Production", "GetNumberSerieSAP");

                oValidValuesSNActive.DoQuery(sNumberSeriesActive);
                oValidValuesSNSAP.DoQuery(sNumberSeriesSAP);

                #endregion

                #region Valores Series de numeración

                if (oValidValuesSNActive.RecordCount > 0)
                {

                    #region Busca las series de numeracion ya parametrizadas

                    oValidValuesSNActive.MoveFirst();

                    for (int K = 0; oValidValuesSNSAP.RecordCount - 1 >= K; K++)
                    {
                        _cboSNPT.ValidValues.Add(oValidValuesSNSAP.Fields.Item("Code_SNPSAP").Value.ToString(), oValidValuesSNSAP.Fields.Item("Name_SNPSAP").Value.ToString());
                        _cboSNPS.ValidValues.Add(oValidValuesSNSAP.Fields.Item("Code_SNPSAP").Value.ToString(), oValidValuesSNSAP.Fields.Item("Name_SNPSAP").Value.ToString());

                        oValidValuesSNSAP.MoveNext();

                    }

                    _cboSNPT.Select(oValidValuesSNActive.Fields.Item("Code_SNPT").Value.ToString(), BoSearchKey.psk_ByValue);
                    _cboSNPS.Select(oValidValuesSNActive.Fields.Item("Code_SNPS").Value.ToString(), BoSearchKey.psk_ByValue);

                    #endregion

                }
                else
                {
                    #region Asigna las series de numeracion de produccion al combo box del formulario

                    oValidValuesSNSAP.MoveFirst();

                    for (int K = 0; oValidValuesSNSAP.RecordCount - 1 >= K; K++)
                    {
                        _cboSNPT.ValidValues.Add(oValidValuesSNSAP.Fields.Item("Code_SNPSAP").Value.ToString(), oValidValuesSNSAP.Fields.Item("Name_SNPSAP").Value.ToString());
                        _cboSNPS.ValidValues.Add(oValidValuesSNSAP.Fields.Item("Code_SNPSAP").Value.ToString(), oValidValuesSNSAP.Fields.Item("Name_SNPSAP").Value.ToString());

                        oValidValuesSNSAP.MoveNext();

                    }

                    #endregion
                }

                DLLFunciones.liberarObjetos(oValidValuesSNSAP);
                DLLFunciones.liberarObjetos(oValidValuesSNActive);

                #endregion

                oFormParProduction.Visible = true;
                oFormParProduction.Refresh();

            }
            catch (Exception e)
            {
                DLLFunciones.sendErrorMessage(sboapp, e);

            }
        }

        public void ChangueFormControlProduction(SAPbouiCOM.Application sboapp, SAPbobsCOM.Company oCompany, SAPbouiCOM.Form oFormControlProduction)
        {
            Funciones.Comunes DLLFunciones = new Funciones.Comunes();

            try
            {
                #region Variables y objetos

                SAPbouiCOM.Matrix oMatrixCOP = (Matrix)oFormControlProduction.Items.Item("MtxCOP").Specific;
                SAPbouiCOM.PictureBox oLogoBO = (SAPbouiCOM.PictureBox)oFormControlProduction.Items.Item("imgLogoBO").Specific;
                SAPbouiCOM.CommonSetting CS = oMatrixCOP.CommonSetting;

                SAPbouiCOM.DataTable oTableCOP = oFormControlProduction.DataSources.DataTables.Add("DT_COP");

                string sConsultaOP;
                int iCount;

                #endregion

                #region Centra en pantalla formulario

                oFormControlProduction.Left = (sboapp.Desktop.Width - oFormControlProduction.Width) / 2;
                oFormControlProduction.Top = (sboapp.Desktop.Height - oFormControlProduction.Height) / 4;

                #endregion

                #region Asignacion Logo BO

                oLogoBO.Picture = (Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Core\\Imagenes\\BO.jpg");

                #endregion

                #region Carga Informacion al Matrix

                sConsultaOP = DLLFunciones.GetStringXMLDocument(oCompany, "BOProduction", "Production", "GetWorkOrders");

                oTableCOP.ExecuteQuery(sConsultaOP);

                iCount = oTableCOP.Rows.Count;

                if (oTableCOP.IsEmpty == false)
                {
                    oMatrixCOP.Clear();

                    oMatrixCOP.Columns.Item("Col_0").DataBind.Bind("DT_COP", "DocNumOPT");

                    oMatrixCOP.Columns.Item("Col_1").DataBind.Bind("DT_COP", "StatusOPT");

                    oMatrixCOP.Columns.Item("Col_2").DataBind.Bind("DT_COP", "ItemCodeOPT");

                    oMatrixCOP.Columns.Item("Col_3").DataBind.Bind("DT_COP", "ItemNameOPT");

                    oMatrixCOP.Columns.Item("Col_4").DataBind.Bind("DT_COP", "WarehouseOPT");

                    oMatrixCOP.Columns.Item("Col_5").DataBind.Bind("DT_COP", "PlannedQtyOPT");

                    oMatrixCOP.Columns.Item("Col_6").DataBind.Bind("DT_COP", "ReceivedQtyOPT");

                    oMatrixCOP.Columns.Item("Col_7").DataBind.Bind("DT_COP", "EtapaProduccion");

                    oMatrixCOP.Columns.Item("Col_8").DataBind.Bind("DT_COP", "ItemCodeOPS");

                    oMatrixCOP.Columns.Item("Col_9").DataBind.Bind("DT_COP", "ItemNameOPS");

                    oMatrixCOP.Columns.Item("Col_10").DataBind.Bind("DT_COP", "WarehouseOPS");

                    oMatrixCOP.Columns.Item("Col_11").DataBind.Bind("DT_COP", "PlannedQtyOPS");

                    oMatrixCOP.Columns.Item("Col_12").DataBind.Bind("DT_COP", "ReceivedQtyOPS");

                    oMatrixCOP.Columns.Item("Col_13").DataBind.Bind("DT_COP", "DocNumOPS");

                    oMatrixCOP.Columns.Item("Col_14").DataBind.Bind("DT_COP", "DocEntryOPS");
                    oMatrixCOP.Columns.Item("Col_14").Visible = false;

                    oMatrixCOP.Columns.Item("Col_15").DataBind.Bind("DT_COP", "imgStatus");

                    oMatrixCOP.Columns.Item("Col_16").DataBind.Bind("DT_COP", "StatusOPS");

                    oMatrixCOP.Columns.Item("Col_17").DataBind.Bind("DT_COP", "QuantityCOLOROPT");
                    oMatrixCOP.Columns.Item("Col_17").Visible = false;

                    oMatrixCOP.Columns.Item("Col_18").DataBind.Bind("DT_COP", "QuantityCOLOROPS");
                    oMatrixCOP.Columns.Item("Col_18").Visible = false;

                    oMatrixCOP.Columns.Item("Col_19").DataBind.Bind("DT_COP", "imgMPDes");

                    oMatrixCOP.Columns.Item("Col_20").DataBind.Bind("DT_COP", "DocEntry");
                    oMatrixCOP.Columns.Item("Col_20").Visible = false;

                    oMatrixCOP.LoadFromDataSource();

                    for (int i = 1; i <= iCount; i++)
                    {
                        oMatrixCOP.CommonSetting.SetCellFontColor(i, 1, DLLFunciones.ColorSB1_MARRON());
                        oMatrixCOP.CommonSetting.SetCellFontColor(i, 2, DLLFunciones.ColorSB1_MARRON());
                        oMatrixCOP.CommonSetting.SetCellFontColor(i, 3, DLLFunciones.ColorSB1_MARRON());
                        oMatrixCOP.CommonSetting.SetCellFontColor(i, 4, DLLFunciones.ColorSB1_MARRON());
                        oMatrixCOP.CommonSetting.SetCellFontColor(i, 5, DLLFunciones.ColorSB1_MARRON());
                        oMatrixCOP.CommonSetting.SetCellFontColor(i, 6, DLLFunciones.ColorSB1_MARRON());
                        oMatrixCOP.CommonSetting.SetCellFontColor(i, 7, DLLFunciones.ColorSB1_MARRON());
                        oMatrixCOP.CommonSetting.SetCellFontColor(i, 16, DLLFunciones.ColorSB1_MARRON());
                        oMatrixCOP.CommonSetting.SetCellFontColor(i, 17, DLLFunciones.ColorSB1_MARRON());

                        if (((SAPbouiCOM.EditText)oMatrixCOP.Columns.Item("Col_16").Cells.Item(i).Specific).Value == "Liberado")
                        {
                            oMatrixCOP.CommonSetting.SetCellBackColor(i, 12, DLLFunciones.ColorSB1_NARANJA());
                        }

                        if (((SAPbouiCOM.EditText)oMatrixCOP.Columns.Item("Col_16").Cells.Item(i).Specific).Value == "Planificado")
                        {
                            oMatrixCOP.CommonSetting.SetCellBackColor(i, 12, DLLFunciones.ColorSB1_AGUA());
                        }

                        if (((SAPbouiCOM.EditText)oMatrixCOP.Columns.Item("Col_16").Cells.Item(i).Specific).Value == "Cerrado")
                        {
                            oMatrixCOP.CommonSetting.SetCellBackColor(i, 12, DLLFunciones.ColorSB1_LIMA());
                        }

                        if (((SAPbouiCOM.EditText)oMatrixCOP.Columns.Item("Col_17").Cells.Item(i).Specific).Value == "VERDE")
                        {
                            oMatrixCOP.CommonSetting.SetCellFontColor(i, 8, DLLFunciones.ColorSB1_VERDE_AZULADO());
                        }
                        else
                        {
                            oMatrixCOP.CommonSetting.SetCellFontColor(i, 8, DLLFunciones.ColorSB1_AZUL());
                        }

                        if (((SAPbouiCOM.EditText)oMatrixCOP.Columns.Item("Col_18").Cells.Item(i).Specific).Value == "VERDE")
                        {
                            oMatrixCOP.CommonSetting.SetCellFontColor(i, 18, DLLFunciones.ColorSB1_VERDE_AZULADO());
                        }
                        else
                        {
                            oMatrixCOP.CommonSetting.SetCellFontColor(i, 18, DLLFunciones.ColorSB1_AZUL());
                        }


                    }

                    oMatrixCOP.AutoResizeColumns();
                }

                #endregion

                oFormControlProduction.Visible = true;
                oFormControlProduction.Refresh();

                DLLFunciones.liberarObjetos(oMatrixCOP);
                DLLFunciones.liberarObjetos(oTableCOP);
                DLLFunciones.liberarObjetos(CS);

            }
            catch (Exception e)
            {
                DLLFunciones.sendErrorMessage(sboapp, e);

            }
        }

        public void UpdateFormControlProduction(SAPbouiCOM.Application sboapp, SAPbobsCOM.Company oCompany, SAPbouiCOM.Form oFormControlProduction)
        {
            Funciones.Comunes DLLFunciones = new Funciones.Comunes();

            try
            {
                #region Variables y objetos

                SAPbouiCOM.Matrix oMatrixCOP = (Matrix)oFormControlProduction.Items.Item("MtxCOP").Specific;
                SAPbouiCOM.CommonSetting CS = oMatrixCOP.CommonSetting;

                SAPbouiCOM.DataTable oTableCOP = oFormControlProduction.DataSources.DataTables.Item("DT_COP");

                string sConsultaOP;
                int iCount;

                #endregion

                oFormControlProduction.Freeze(true);

                #region Carga Informacion al Matrix

                sConsultaOP = DLLFunciones.GetStringXMLDocument(oCompany, "BOProduction", "Production", "GetWorkOrders");

                oTableCOP.ExecuteQuery(sConsultaOP);

                iCount = oTableCOP.Rows.Count;

                if (iCount > 0)
                {
                    oMatrixCOP.Clear();

                    oMatrixCOP.Columns.Item("Col_0").DataBind.Bind("DT_COP", "DocNumOPT");

                    oMatrixCOP.Columns.Item("Col_1").DataBind.Bind("DT_COP", "StatusOPT");

                    oMatrixCOP.Columns.Item("Col_2").DataBind.Bind("DT_COP", "ItemCodeOPT");

                    oMatrixCOP.Columns.Item("Col_3").DataBind.Bind("DT_COP", "ItemNameOPT");

                    oMatrixCOP.Columns.Item("Col_4").DataBind.Bind("DT_COP", "WarehouseOPT");

                    oMatrixCOP.Columns.Item("Col_5").DataBind.Bind("DT_COP", "PlannedQtyOPT");

                    oMatrixCOP.Columns.Item("Col_6").DataBind.Bind("DT_COP", "ReceivedQtyOPT");

                    oMatrixCOP.Columns.Item("Col_7").DataBind.Bind("DT_COP", "EtapaProduccion");

                    oMatrixCOP.Columns.Item("Col_8").DataBind.Bind("DT_COP", "ItemCodeOPS");

                    oMatrixCOP.Columns.Item("Col_9").DataBind.Bind("DT_COP", "ItemNameOPS");

                    oMatrixCOP.Columns.Item("Col_10").DataBind.Bind("DT_COP", "WarehouseOPS");

                    oMatrixCOP.Columns.Item("Col_11").DataBind.Bind("DT_COP", "PlannedQtyOPS");

                    oMatrixCOP.Columns.Item("Col_12").DataBind.Bind("DT_COP", "ReceivedQtyOPS");

                    oMatrixCOP.Columns.Item("Col_13").DataBind.Bind("DT_COP", "DocNumOPS");

                    oMatrixCOP.Columns.Item("Col_14").DataBind.Bind("DT_COP", "DocEntryOPS");
                    oMatrixCOP.Columns.Item("Col_14").Visible = false;

                    oMatrixCOP.Columns.Item("Col_15").DataBind.Bind("DT_COP", "imgStatus");

                    oMatrixCOP.Columns.Item("Col_16").DataBind.Bind("DT_COP", "StatusOPS");

                    oMatrixCOP.LoadFromDataSource();

                    for (int i = 1; i <= iCount; i++)
                    {
                        oMatrixCOP.CommonSetting.SetCellFontColor(i, 1, DLLFunciones.ColorSB1_MARRON());
                        oMatrixCOP.CommonSetting.SetCellFontColor(i, 2, DLLFunciones.ColorSB1_MARRON());
                        oMatrixCOP.CommonSetting.SetCellFontColor(i, 3, DLLFunciones.ColorSB1_MARRON());
                        oMatrixCOP.CommonSetting.SetCellFontColor(i, 4, DLLFunciones.ColorSB1_MARRON());
                        oMatrixCOP.CommonSetting.SetCellFontColor(i, 5, DLLFunciones.ColorSB1_MARRON());
                        oMatrixCOP.CommonSetting.SetCellFontColor(i, 6, DLLFunciones.ColorSB1_MARRON());
                        oMatrixCOP.CommonSetting.SetCellFontColor(i, 7, DLLFunciones.ColorSB1_AZUL());
                        oMatrixCOP.CommonSetting.SetCellFontColor(i, 16, DLLFunciones.ColorSB1_AZUL());

                        if (((SAPbouiCOM.EditText)oMatrixCOP.Columns.Item("Col_16").Cells.Item(i).Specific).Value == "Liberado")
                        {
                            oMatrixCOP.CommonSetting.SetCellBackColor(i, 10, DLLFunciones.ColorSB1_NARANJA());
                        }

                        if (((SAPbouiCOM.EditText)oMatrixCOP.Columns.Item("Col_16").Cells.Item(i).Specific).Value == "Planificado")
                        {
                            oMatrixCOP.CommonSetting.SetCellBackColor(i, 10, DLLFunciones.ColorSB1_AGUA());
                        }

                        if (((SAPbouiCOM.EditText)oMatrixCOP.Columns.Item("Col_16").Cells.Item(i).Specific).Value == "Cerrado")
                        {
                            oMatrixCOP.CommonSetting.SetCellBackColor(i, 10, DLLFunciones.ColorSB1_LIMA());
                        }
                    }

                    oMatrixCOP.AutoResizeColumns();
                }

                #endregion

                oFormControlProduction.Freeze(false);
                oFormControlProduction.Refresh();

                DLLFunciones.liberarObjetos(oMatrixCOP);
                DLLFunciones.liberarObjetos(oTableCOP);
                DLLFunciones.liberarObjetos(CS);

            }
            catch (Exception e)
            {
                DLLFunciones.sendErrorMessage(sboapp, e);

            }
        }

        public void LoadFormMRawMaterial(SAPbouiCOM.Application sboapp, SAPbobsCOM.Company oCompany, SAPbouiCOM.Form oFormControlProduction, ItemEvent pVal)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            try
            {
                #region Variables y objetos

                SAPbouiCOM.Matrix oMatrixCOP = (Matrix)oFormControlProduction.Items.Item("MtxCOP").Specific;
                SAPbobsCOM.Recordset oRsCOE = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string sConsultaMPE;
                string sDocEntryOPS;
                int iCount;

                sDocEntryOPS = ((SAPbouiCOM.EditText)oMatrixCOP.Columns.Item("Col_14").Cells.Item(pVal.Row).Specific).Value;

                #endregion

                #region Consulta si existe materia prima entregada

                sConsultaMPE = DllFunciones.GetStringXMLDocument(oCompany, "BOProduction", "Production", "GetMPE");

                sConsultaMPE = sConsultaMPE.Replace("%DocEntryOPS%", sDocEntryOPS);

                oRsCOE.DoQuery(sConsultaMPE);

                iCount = oRsCOE.RecordCount;

                DllFunciones.liberarObjetos(oRsCOE);

                #endregion

                if (iCount > 0)
                {
                    #region Carga Formulario Materia Prima Entregada

                    string ArchivoSRF = "materia_prima_entregada.srf";
                    DllFunciones.LoadFromXML(sboapp, "BOProduction", ref ArchivoSRF);

                    SAPbouiCOM.Form oFormRawMaterial;
                    oFormRawMaterial = sboapp.Forms.Item("BOFormMPC");

                    #endregion

                    #region Centra en pantalla formulario

                    oFormRawMaterial.Left = (sboapp.Desktop.Width - oFormRawMaterial.Width) / 2;
                    oFormRawMaterial.Top = (sboapp.Desktop.Height - oFormRawMaterial.Height) / 4;

                    #endregion

                    #region Consulta infomacion matertia prima entregada 

                    SAPbouiCOM.DataTable oTableMPE = oFormRawMaterial.DataSources.DataTables.Add("DT_MPE");
                    oTableMPE.ExecuteQuery(sConsultaMPE);

                    #endregion

                    oFormRawMaterial.Freeze(true);

                    #region Carga Informacion al Matrix

                    SAPbouiCOM.Matrix oMatrixMPE = (Matrix)oFormRawMaterial.Items.Item("MtxMPE").Specific;

                    oMatrixMPE.Clear();

                    oMatrixMPE.Columns.Item("Col_0").DataBind.Bind("DT_MPE", "DocEntry");
                    oMatrixMPE.Columns.Item("Col_0").Visible = false;

                    oMatrixMPE.Columns.Item("Col_1").DataBind.Bind("DT_MPE", "DocNum");

                    oMatrixMPE.Columns.Item("Col_2").DataBind.Bind("DT_MPE", "DocDate");

                    oMatrixMPE.Columns.Item("Col_3").DataBind.Bind("DT_MPE", "ItemCode");

                    oMatrixMPE.Columns.Item("Col_4").DataBind.Bind("DT_MPE", "Dscription");

                    oMatrixMPE.Columns.Item("Col_5").DataBind.Bind("DT_MPE", "Quantity");

                    oMatrixMPE.Columns.Item("Col_6").DataBind.Bind("DT_MPE", "WhsCode");

                    oMatrixMPE.Columns.Item("Col_7").DataBind.Bind("DT_MPE", "OF");

                    oMatrixMPE.LoadFromDataSource();

                    oMatrixMPE.AutoResizeColumns();

                    oFormRawMaterial.Visible = true;
                    oFormRawMaterial.Freeze(false);
                    oFormRawMaterial.Refresh();

                    DllFunciones.liberarObjetos(oMatrixMPE);
                    DllFunciones.liberarObjetos(oTableMPE);
                }

                #endregion

                DllFunciones.liberarObjetos(oMatrixCOP);
            }
            catch (Exception e)
            {
                DllFunciones.sendErrorMessage(sboapp, e);

            }
        }

        public void AddItemsToDocumets(SAPbouiCOM.Form oFormProduction)
        {
            #region Variables y objetos 

            SAPbouiCOM.ComboBox oTO = null;
            SAPbouiCOM.Item oUDFProduction = null;
            SAPbouiCOM.Item oUDF = null;
            SAPbouiCOM.Item oDataMasterDate = null;
            SAPbouiCOM.StaticText oStaticText = null;

            oUDF = oFormProduction.Items.Item("78");
            oDataMasterDate = oFormProduction.Items.Item("6");

            #endregion

            #region Campo Tipo de Orden

            //*******************************************
            // Se adiciona Label "Tipo de Orden"
            //*******************************************

            oUDFProduction = oFormProduction.Items.Add("lblTO", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oUDFProduction.Left = oUDF.Left + 130;
            oUDFProduction.Width = oUDF.Width - 50;
            oUDFProduction.Top = oUDF.Top;
            oUDFProduction.Height = oUDF.Height;

            oUDFProduction.LinkTo = "txtTO";

            oStaticText = ((SAPbouiCOM.StaticText)(oUDFProduction.Specific));

            oStaticText.Caption = "Tipo de Orden";

            //*******************************************
            // Se adiciona Tex Box "Tipo de Orden"
            //*******************************************

            oUDFProduction = oFormProduction.Items.Add("txtTO", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            oUDFProduction.Left = oUDF.Left + 205;
            oUDFProduction.Width = oUDF.Width;
            oUDFProduction.Top = oUDF.Top;
            oUDFProduction.Height = oUDF.Height;
            //oUDFProduction.Enabled = false;

            oUDFProduction.DisplayDesc = true;

            oFormProduction.DataSources.UserDataSources.Add("cboTO", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);

            oTO = ((SAPbouiCOM.ComboBox)(oUDFProduction.Specific));


            oTO.DataBind.SetBound(true, "OWOR", "U_BO_TO");


            oTO.ValidValues.Add("T", "Producto Terminado");
            oTO.ValidValues.Add("S", "Prodcuto Semielaborado");

            if (oFormProduction.Mode == BoFormMode.fm_ADD_MODE)
            {
                oTO.Select("T", BoSearchKey.psk_ByValue);
            }

            #endregion

            oDataMasterDate.Click();

        }

        public void UpdateParametersProduction(SAPbouiCOM.Application sboapp, SAPbobsCOM.Company _oCompany, SAPbouiCOM.Form _oFormParProduction)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            #region Variables y objetos

            SAPbobsCOM.UserTable oUDT = (SAPbobsCOM.UserTable)(_oCompany.UserTables.Item("BOPRODP"));

            SAPbouiCOM.ComboBox cboSNPT = (ComboBox)_oFormParProduction.Items.Item("txtSNPT").Specific;
            SAPbouiCOM.ComboBox cboSNPS = (ComboBox)_oFormParProduction.Items.Item("txtSNPS").Specific;
            SAPbouiCOM.Button btnOK = (Button)_oFormParProduction.Items.Item("btnUpdate").Specific;

            string sValidateParametersProduccion = null;

            SAPbobsCOM.Recordset oParametersProduccion = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            #endregion

            #region Consulta si existe parametros configurados de produccion

            sValidateParametersProduccion = DllFunciones.GetStringXMLDocument(_oCompany, "BOProduction", "Production", "GetParametersProduction");

            oParametersProduccion.DoQuery(sValidateParametersProduccion);

            #endregion

            if (oParametersProduccion.RecordCount > 0)
            {
                #region Si existe, actualice el code 

                oUDT.GetByKey(Convert.ToString(oParametersProduccion.Fields.Item("Code").Value.ToString()));
                oUDT.UserFields.Fields.Item("U_BO_SNPT").Value = cboSNPT.Selected.Value;
                oUDT.UserFields.Fields.Item("U_BO_SNPS").Value = cboSNPS.Selected.Value;

                oUDT.Update();

                #endregion
            }
            else
            {
                #region Variables y Objetos

                string sSearchNextCode = null;

                SAPbobsCOM.Recordset oSearchNextCode = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                #endregion

                #region Consulta el Code a asignar

                sSearchNextCode = DllFunciones.GetStringXMLDocument(_oCompany, "BOProduction", "Production", "SerachNextCode");

                oSearchNextCode.DoQuery(sSearchNextCode);

                #endregion

                #region Si no existe, inserta el code 

                oUDT.Code = Convert.ToString(oSearchNextCode.Fields.Item("ID").Value.ToString());
                oUDT.Name = Convert.ToString(oSearchNextCode.Fields.Item("ID").Value.ToString());
                oUDT.UserFields.Fields.Item("U_BO_SNPT").Value = cboSNPT.Selected.Value;
                oUDT.UserFields.Fields.Item("U_BO_SNPS").Value = cboSNPS.Selected.Value;

                oUDT.Add();

                #endregion

            }

            DllFunciones.StatusBar(sboapp, BoStatusBarMessageType.smt_Success, "Actualizado correctamente...");
            btnOK.Caption = "OK";
            _oFormParProduction.Mode = BoFormMode.fm_OK_MODE;
            _oFormParProduction.Refresh();


        }

        public void LinkedButtonMatrixFormCOP(SAPbouiCOM.Application _sboapp, SAPbobsCOM.Company _oCompany, SAPbouiCOM.Form oFormCOP, ItemEvent pVal, string sColumna)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            #region Consulta DocEntry Orden de Produccion

            SAPbouiCOM.Matrix oMatrixCOP = (Matrix)oFormCOP.Items.Item("MtxCOP").Specific;

            #endregion

            if (sColumna == "Col_0")
            {
                #region LinkeButton Orden de produccion producto terminado 

                ((EditText)oFormCOP.Items.Item("txtValor").Specific).Value = ((SAPbouiCOM.EditText)oMatrixCOP.Columns.Item("Col_20").Cells.Item(pVal.Row).Specific).Value;
                Item itm = oFormCOP.Items.Item("lbValor");
                ((LinkedButton)itm.Specific).LinkedObjectType = "202";
                itm.Click();

                #endregion
            }
            else if (sColumna == "Col_13")
            {
                #region LinkeButton Orden de produccion producto semielaborado

                ((EditText)oFormCOP.Items.Item("txtValor").Specific).Value = ((SAPbouiCOM.EditText)oMatrixCOP.Columns.Item("Col_14").Cells.Item(pVal.Row).Specific).Value;
                Item itm = oFormCOP.Items.Item("lbValor");
                ((LinkedButton)itm.Specific).LinkedObjectType = "202";
                itm.Click();

                #endregion
            }
            else if (sColumna == "Col_1")
            {
                #region LinkeButton abre articulo Producto terminado

                oFormCOP.Freeze(true);

                SAPbouiCOM.Column oColumDocEntry;
                SAPbouiCOM.LinkedButton LkBtnDocEntry;

                oColumDocEntry = oMatrixCOP.Columns.Item("Col_2");

                LkBtnDocEntry = (SAPbouiCOM.LinkedButton)oColumDocEntry.ExtendedObject;
                LkBtnDocEntry.LinkedObject = BoLinkedObject.lf_Items;

                oFormCOP.Freeze(false);

                #endregion
            }
            else if (sColumna == "Col_8")
            {
                #region LinkeButton abre Producto Semielaborado

                oFormCOP.Freeze(true);

                SAPbouiCOM.Column oColumDocEntry;
                SAPbouiCOM.LinkedButton LkBtnDocEntry;

                oColumDocEntry = oMatrixCOP.Columns.Item("4");

                LkBtnDocEntry = (SAPbouiCOM.LinkedButton)oColumDocEntry.ExtendedObject;
                LkBtnDocEntry.LinkedObject = BoLinkedObject.lf_Items;

                oFormCOP.Freeze(false);

                #endregion
            }
        }

        public void LinkedButtonMatrixFormMPE(SAPbouiCOM.Application _sboapp, SAPbobsCOM.Company _oCompany, SAPbouiCOM.Form BOFormMPC, ItemEvent pVal, string sColumna)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            #region Consulta DocEntry Orden de Produccion

            SAPbouiCOM.Matrix oMatrixMPE = (Matrix)BOFormMPC.Items.Item("MtxMPE").Specific;

            #endregion

            if (sColumna == "Col_1")
            {
                #region LinkeButton Materia prima consumida 

                ((EditText)BOFormMPC.Items.Item("txtValor").Specific).Value = ((SAPbouiCOM.EditText)oMatrixMPE.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific).Value;
                string doceentry = ((SAPbouiCOM.EditText)oMatrixMPE.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific).Value;
                Item itm = BOFormMPC.Items.Item("lbValor");
                ((LinkedButton)itm.Specific).LinkedObjectType = "60";
                itm.Click();

                #endregion
            }
        }

        public string VersionDll()
        {
            try
            {

                Assembly Assembly = Assembly.LoadFrom("BOProduccion.dll");
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

