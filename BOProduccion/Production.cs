using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbobsCOM;
using SAPbouiCOM;
using Funciones;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using System.Globalization;


namespace BOProduccion
{
    public class Production
    {
        public void AddItemsToWorkOrder(SAPbouiCOM.Form _oFormWorkOrder)
        {
            #region Variables y objetos 

            SAPbouiCOM.ComboBox oTO = null;
            SAPbouiCOM.Item oUDFProduction = null;
            SAPbouiCOM.Item oUDF = null;
            SAPbouiCOM.Item oDataMasterDate = null;
            SAPbouiCOM.StaticText oStaticText = null;

            oUDF = _oFormWorkOrder.Items.Item("78");
            oDataMasterDate = _oFormWorkOrder.Items.Item("6");

            #endregion

            #region Campo Tipo de Orden

            //*******************************************
            // Se adiciona Label "Tipo de Orden"
            //*******************************************

            oUDFProduction = _oFormWorkOrder.Items.Add("lblTO", SAPbouiCOM.BoFormItemTypes.it_STATIC);
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

            oUDFProduction = _oFormWorkOrder.Items.Add("txtTO", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            oUDFProduction.Left = oUDF.Left + 205;
            oUDFProduction.Width = oUDF.Width;
            oUDFProduction.Top = oUDF.Top;
            oUDFProduction.Height = oUDF.Height;
            oUDFProduction.Enabled = false;

            oUDFProduction.DisplayDesc = true;

            _oFormWorkOrder.DataSources.UserDataSources.Add("cboTO", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);

            oTO = ((SAPbouiCOM.ComboBox)(oUDFProduction.Specific));


            oTO.DataBind.SetBound(true, "OWOR", "U_BO_TO");


            oTO.ValidValues.Add("T", "Producto Terminado");
            oTO.ValidValues.Add("S", "Prodcuto Semielaborado");

            if (_oFormWorkOrder.Mode == BoFormMode.fm_ADD_MODE)
            {
                oTO.Select("T", BoSearchKey.psk_ByValue);
            }

            #endregion

            #region Adicion Panel Ruta de produccion

            //SAPbouiCOM.Form _oFormWorOrder;
            //SAPbouiCOM.Item _oNewItem;
            //SAPbouiCOM.Item _oItem;
            //SAPbouiCOM.Folder _oFolderItem;

            //_oFormWorOrder = _oFormWorkOrder;
            //_oNewItem = _oFormWorOrder.Items.Add("FolderBO1", SAPbouiCOM.BoFormItemTypes.it_FOLDER);
            //_oItem = _oFormWorOrder.Items.Item("234000008");

            //_oNewItem.Top = _oItem.Top;
            //_oNewItem.Height = _oItem.Height;
            //_oNewItem.Width = _oItem.Width;
            //_oNewItem.Left = _oItem.Left + _oItem.Width;

            //_oFolderItem = ((SAPbouiCOM.Folder)(_oNewItem.Specific));

            //_oFolderItem.Caption = "Ruta de Producción";

            //_oFolderItem.GroupWith("234000008");

            ////ItemsDocuments(_oFormInvoices, _TipoDoc);

            //_oFormWorOrder.PaneLevel = 1;

            #endregion

            #region Adicionar Matrix Matrix Ruta de produccion

            //AddMatrixToFormWorkOrderRouteProduction(_oFormWorkOrder);

            #endregion

            oDataMasterDate.Click();

        }

        private void AddChooseFromListoOITM(SAPbouiCOM.Application _sboapp, SAPbouiCOM.Form _oFormWO)
        {

            SAPbouiCOM.ChooseFromListCollection oCFLs = null;
            SAPbouiCOM.Conditions oCons = null;
            SAPbouiCOM.Condition oCon = null;

            oCFLs = _oFormWO.ChooseFromLists;

            SAPbouiCOM.ChooseFromList oCFL = null;
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
            oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(_sboapp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)));

            //  Adding 2 CFL, one for the button and one for the edit text.
            oCFLCreationParams.MultiSelection = false;
            oCFLCreationParams.ObjectType = "4";
            oCFLCreationParams.UniqueID = "CFL1";

            oCFL = oCFLs.Add(oCFLCreationParams);

            //  Adding Conditions to CFL1

            oCons = oCFL.GetConditions();

            oCon = oCons.Add();
            oCon.Alias = "TreeType";
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCon.CondVal = "P";
            oCFL.SetConditions(oCons);

            oCFLCreationParams.UniqueID = "CFL2";
            oCFL = oCFLs.Add(oCFLCreationParams);
        }

        private void AddChooseFromListoOITMMatrix(SAPbouiCOM.Application _sboapp, SAPbouiCOM.Form _oFormWO)
        {

            SAPbouiCOM.ChooseFromListCollection oCFLs = null;
            SAPbouiCOM.Conditions oCons = null;
            SAPbouiCOM.Condition oCon = null;

            oCFLs = _oFormWO.ChooseFromLists;

            SAPbouiCOM.ChooseFromList oCFL = null;
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
            oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(_sboapp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)));

            //  Adding 2 CFL, one for the button and one for the edit text.
            oCFLCreationParams.MultiSelection = false;
            oCFLCreationParams.ObjectType = "4";
            oCFLCreationParams.UniqueID = "CFL3";

            oCFL = oCFLs.Add(oCFLCreationParams);

            //  Adding Conditions to CFL1

            oCons = oCFL.GetConditions();

            oCon = oCons.Add();
            oCon.Alias = "TreeType";
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCon.CondVal = "P";
            oCFL.SetConditions(oCons);

            oCFLCreationParams.UniqueID = "CFL4";
            oCFL = oCFLs.Add(oCFLCreationParams);
        }

        private void AddChooseFromListAccount(SAPbouiCOM.Application _sboapp, SAPbouiCOM.Form _oFormWO)
        {

            SAPbouiCOM.ChooseFromListCollection oCFLs = null;
            SAPbouiCOM.Conditions oCons = null;
            SAPbouiCOM.Condition oCon = null;

            oCFLs = _oFormWO.ChooseFromLists;

            SAPbouiCOM.ChooseFromList oCFL = null;
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
            oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(_sboapp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)));

            //  Adding 2 CFL, one for the button and one for the edit text.
            oCFLCreationParams.MultiSelection = false;
            oCFLCreationParams.ObjectType = "1";
            oCFLCreationParams.UniqueID = "CFL1";

            oCFL = oCFLs.Add(oCFLCreationParams);

            //  Adding Conditions to CFL1

            oCons = oCFL.GetConditions();

            oCon = oCons.Add();
            oCon.Alias = "Postable";
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCon.CondVal = "Y";
            oCFL.SetConditions(oCons);

            oCFLCreationParams.UniqueID = "CFL2";
            oCFL = oCFLs.Add(oCFLCreationParams);
        }

        private void AddCFLOITMBatchNumber(SAPbouiCOM.Application _sboapp, SAPbouiCOM.Form _oFormGL)
        {

            SAPbouiCOM.ChooseFromListCollection oCFLs = null;
            SAPbouiCOM.Conditions oCons = null;
            SAPbouiCOM.Condition oCon = null;

            oCFLs = _oFormGL.ChooseFromLists;

            SAPbouiCOM.ChooseFromList oCFL = null;
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
            oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(_sboapp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)));

            //  Adding 2 CFL, one for the button and one for the edit text.
            oCFLCreationParams.MultiSelection = false;
            oCFLCreationParams.ObjectType = "4";
            oCFLCreationParams.UniqueID = "CFL1";

            oCFL = oCFLs.Add(oCFLCreationParams);

            //  Adding Conditions to CFL1

            oCons = oCFL.GetConditions();

            oCon = oCons.Add();
            oCon.Alias = "ManBtchNum";
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCon.CondVal = "Y";
            oCFL.SetConditions(oCons);

            oCFLCreationParams.UniqueID = "CFL2";
            oCFL = oCFLs.Add(oCFLCreationParams);
        }

        private void AddCFLWareHouse(SAPbouiCOM.Application _sboapp, SAPbouiCOM.Form _oFormGL)
        {

            SAPbouiCOM.ChooseFromListCollection oCFLs = null;
            SAPbouiCOM.Conditions oCons = null;
            SAPbouiCOM.Condition oCon = null;

            oCFLs = _oFormGL.ChooseFromLists;

            SAPbouiCOM.ChooseFromList oCFL = null;
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
            oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(_sboapp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)));

            //  Adding 2 CFL, one for the button and one for the edit text.
            oCFLCreationParams.MultiSelection = false;
            oCFLCreationParams.ObjectType = "64";
            oCFLCreationParams.UniqueID = "CFL3";

            oCFL = oCFLs.Add(oCFLCreationParams);

            //  Adding Conditions to CFL1

            oCons = oCFL.GetConditions();

            oCon = oCons.Add();
            oCon.Alias = "Inactive";
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCon.CondVal = "N";
            oCFL.SetConditions(oCons);

            oCFLCreationParams.UniqueID = "CFL4";
            oCFL = oCFLs.Add(oCFLCreationParams);
        }

        private void AddCFLProductionRoute(SAPbouiCOM.Application _sboapp, SAPbouiCOM.Form _oFormWO)
        {

            SAPbouiCOM.ChooseFromListCollection oCFLs = null;

            oCFLs = _oFormWO.ChooseFromLists;

            SAPbouiCOM.ChooseFromList oCFL = null;
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
            oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(_sboapp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)));

            //  Adding 2 CFL, one for the button and one for the edit text.
            oCFLCreationParams.MultiSelection = false;
            oCFLCreationParams.ObjectType = "BORP";
            oCFLCreationParams.UniqueID = "CFL3";

            oCFL = oCFLs.Add(oCFLCreationParams);

        }

        private void AddCFLProductionRouteNWO(SAPbouiCOM.Application _sboapp, SAPbouiCOM.Form _oFormWO)
        {

            SAPbouiCOM.ChooseFromListCollection oCFLs = null;

            oCFLs = _oFormWO.ChooseFromLists;

            SAPbouiCOM.ChooseFromList oCFL = null;
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
            oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(_sboapp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)));

            //  Adding 2 CFL, one for the button and one for the edit text.
            oCFLCreationParams.MultiSelection = false;
            oCFLCreationParams.ObjectType = "BORP";
            oCFLCreationParams.UniqueID = "CFL1";

            oCFL = oCFLs.Add(oCFLCreationParams);

        }

        public void AddNewRowMatrix(SAPbouiCOM.Form oFormNewWO)
        {
            #region Variables y Objetos

            int Counter = 0;

            SAPbouiCOM.Matrix oMatrixNWO = (SAPbouiCOM.Matrix)oFormNewWO.Items.Item("mtxRP").Specific;

            Counter = oMatrixNWO.RowCount + 1;

            #endregion

            #region Limipia datasource

            oFormNewWO.DataSources.UserDataSources.Item("DSCol0").ValueEx = null;
            oFormNewWO.DataSources.UserDataSources.Item("DSCol1").ValueEx = null;
            oFormNewWO.DataSources.UserDataSources.Item("DSCol2").ValueEx = null;

            #endregion

            oMatrixNWO.AddRow();

            oMatrixNWO.Columns.Item("Col_1").Cells.Item(Counter).Click();

        }

        public void AddLineMatrixLO(SAPbouiCOM.Form oFormGL, ItemEvent pVal)
        {
            #region Variables y Objetos

            int Counter = 0;

            SAPbouiCOM.Matrix oMatrixLO = (SAPbouiCOM.Matrix)oFormGL.Items.Item("mtxLO").Specific;

            Counter = oMatrixLO.RowCount + 1;

            #endregion

            #region Limipia datasource

            oFormGL.DataSources.UserDataSources.Item("DSCol0").ValueEx = null;
            oFormGL.DataSources.UserDataSources.Item("DSCol1").ValueEx = "0";
            oFormGL.DataSources.UserDataSources.Item("DSCol2").ValueEx = "0";

            #endregion

            oMatrixLO.AddRow();

            oMatrixLO.SetCellFocus(Counter, 1);

            oMatrixLO.FlushToDataSource();

        }

        public void AddLineMatrixLD(SAPbouiCOM.Form oFormGL, ItemEvent pVal)
        {
            #region Variables y Objetos

            int Counter = 0;

            SAPbouiCOM.Matrix oMatrixLD = (SAPbouiCOM.Matrix)oFormGL.Items.Item("mtxLD").Specific;

            Counter = oMatrixLD.RowCount + 1;

            #endregion

            #region Limipia datasource

            oFormGL.DataSources.UserDataSources.Item("DSCol4").ValueEx = null;
            oFormGL.DataSources.UserDataSources.Item("DSCol6").ValueEx = "0";

            #endregion

            oMatrixLD.AddRow();

            oMatrixLD.SetCellFocus(Counter, 1);

        }

        public void DeleteRowMatrix(SAPbouiCOM.Application _sboapp, SAPbouiCOM.Form oFormNewWO)
        {
            #region Variables y Objetos

            int RowIndex = 0;

            SAPbouiCOM.Matrix oMatrixNWO = (SAPbouiCOM.Matrix)oFormNewWO.Items.Item("mtxRP").Specific;

            RowIndex = oMatrixNWO.GetNextSelectedRow();

            #endregion

            #region Valida que no se pueda eliminar el producto terminado

            if (RowIndex == 1)
            {
                Funciones.Comunes DllFunciones = new Funciones.Comunes();

                DllFunciones.sendMessageBox(_sboapp, "No se puede eliminar la linea 1, corresponde al producto terminado");

            }
            else
            {
                oMatrixNWO.DeleteRow(RowIndex);
            }

            #endregion
        }

        public void DeleteRowMatrixLO(SAPbouiCOM.Application _sboapp, SAPbouiCOM.Form oFormGL, ItemEvent pVal)
        {
            try
            {
                #region Variables y Objetos

                int RowIndex = 0;

                SAPbouiCOM.Matrix oMatrixLO = (SAPbouiCOM.Matrix)oFormGL.Items.Item("mtxLO").Specific;

                var ValidateRowIndex = oMatrixLO.GetCellFocus();

                #endregion

                if (ValidateRowIndex != null)
                {
                    #region Obtiene el valor de la celda

                    SAPbouiCOM.CellPosition _Cell;

                    _Cell = oMatrixLO.GetCellFocus();

                    RowIndex = _Cell.rowIndex;

                    #endregion

                    #region Valida que no se pueda eliminar el producto terminado

                    if (RowIndex == 1)
                    {
                        Funciones.Comunes DllFunciones = new Funciones.Comunes();

                        DllFunciones.sendMessageBox(_sboapp, "No se puede eliminar la linea 1, se nesecita al menos 1 lote para transferir");
                    }
                    else
                    {
                        oMatrixLO.DeleteRow(RowIndex);
                    }

                    #endregion
                }
            }
            catch (Exception)
            {

                throw;
            }
        }

        public void DeleteRowMatrixLD(SAPbouiCOM.Application _sboapp, SAPbouiCOM.Form oFormGL, ItemEvent pVal)
        {
            try
            {
                #region Variables y Objetos

                int RowIndex = 0;

                SAPbouiCOM.Matrix oMatrixLD = (SAPbouiCOM.Matrix)oFormGL.Items.Item("mtxLD").Specific;

                var ValidateRowIndex = oMatrixLD.GetCellFocus();

                #endregion

                if (ValidateRowIndex != null)
                {
                    #region Obtiene el valor de la celda

                    SAPbouiCOM.CellPosition _Cell;

                    _Cell = oMatrixLD.GetCellFocus();

                    RowIndex = _Cell.rowIndex;

                    #endregion

                    #region Valida que no se pueda eliminar el producto terminado

                    if (RowIndex == 1)
                    {
                        Funciones.Comunes DllFunciones = new Funciones.Comunes();

                        DllFunciones.sendMessageBox(_sboapp, "No se puede eliminar la linea 1, se nesecita al menos 1 lote para transferir");
                    }
                    else
                    {
                        oMatrixLD.DeleteRow(RowIndex);
                    }

                    #endregion
                }
            }
            catch (Exception)
            {

                throw;
            }
        }

        public void CreateUDTandUDFProduction(SAPbouiCOM.Application sboapp, SAPbobsCOM.Company oCompany)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            #region Creacion de tablas

            DllFunciones.ProgressBar(oCompany, sboapp, 32, 1, "Creando Tabla - Parametros Produccion Avanzada, por favor espere...");
            DllFunciones.crearTabla(oCompany, sboapp, "BOPRODP", "BO-Param. Produc. Avan.", SAPbobsCOM.BoUTBTableType.bott_NoObject);

            DllFunciones.ProgressBar(oCompany, sboapp, 32, 1, "Creando Tabla - BORTDC - BO Registro de tiempo detallado , por favor espere...");
            DllFunciones.crearTabla(oCompany, sboapp, "BORTDC", "BO-Reg.Tiem.Deta.Ca", SAPbobsCOM.BoUTBTableType.bott_Document);

            DllFunciones.ProgressBar(oCompany, sboapp, 32, 1, "Creando Tabla - BORTDD - BO Registro de tiempo detallado , por favor espere...");
            DllFunciones.crearTabla(oCompany, sboapp, "BORTDD", "BO-Reg.Tiem.Deta.De", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);

            DllFunciones.ProgressBar(oCompany, sboapp, 32, 1, "Creando Tabla - BOTL - BO Transferencia de lotes , por favor espere...");
            DllFunciones.crearTabla(oCompany, sboapp, "BOTL", "BO Transferencia lotes", SAPbobsCOM.BoUTBTableType.bott_NoObject);

            DllFunciones.ProgressBar(oCompany, sboapp, 32, 1, "Creando Tabla - BORP - BO Ruta de producción encabezado, por favor espere...");
            DllFunciones.crearTabla(oCompany, sboapp, "BORPEN", "BO Ruta Prod. Encabezado", SAPbobsCOM.BoUTBTableType.bott_MasterData);

            DllFunciones.ProgressBar(oCompany, sboapp, 32, 1, "Creando Tabla - BORP - BO Ruta de producción lineas, por favor espere...");
            DllFunciones.crearTabla(oCompany, sboapp, "BORPLI", "BO Ruta Prod. Lineas", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);

            #endregion

            #region Creacion de Campos

            DllFunciones.ProgressBar(oCompany, sboapp, 32, 1, "Creando Campo - OWOR - Tipo de Orden ...");
            string[] ValidValuesFields1 = { "T", "Prodcuto Terminado", "S", "Producto Semielaborado" };
            DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, "", BoYesNoEnum.tNO, ValidValuesFields1, "OWOR", "BO_TO", "Tipo Orden");

            DllFunciones.ProgressBar(oCompany, sboapp, 32, 1, "Creando Campo - OWOR - OP Principal.. ");
            DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 30, "", BoYesNoEnum.tNO, null, "OWOR", "BO_OPP", "OP Principal");

            DllFunciones.ProgressBar(oCompany, sboapp, 32, 1, "Creando Campo - OWOR - Posicion Articulo.. ");
            DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 30, "", BoYesNoEnum.tNO, null, "OWOR", "BO_PosId", "Posicion OP");

            DllFunciones.ProgressBar(oCompany, sboapp, 32, 1, "Creando Campo - BOPRODP - Serie numeracion Produc. Terminado... ");
            DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 30, "", BoYesNoEnum.tNO, null, "@BOPRODP", "BO_SNPT", "Ser.Num.PP");

            DllFunciones.ProgressBar(oCompany, sboapp, 32, 1, "Creando Campo - BOPRODP - Serie numeracion Produc. Semielaborado... ");
            DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 30, "", BoYesNoEnum.tNO, null, "@BOPRODP", "BO_SNPS", "Ser.Num.PS");

            DllFunciones.ProgressBar(oCompany, sboapp, 32, 1, "Creando Campo - BOPRODP - Serie numeracion Salida de Mercancia... ");
            DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 30, "", BoYesNoEnum.tNO, null, "@BOPRODP", "BO_SNSM", "Ser.Num.SM");

            DllFunciones.ProgressBar(oCompany, sboapp, 32, 1, "Creando Campo - BOPRODP - Serie numeracion Entrada de Mercancia... ");
            DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 30, "", BoYesNoEnum.tNO, null, "@BOPRODP", "BO_SNEM", "Ser.Num.EM");

            DllFunciones.ProgressBar(oCompany, sboapp, 32, 1, "Creando Campo - BOPRODP - Ruta Imagenes ... ");
            DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "@BOPRODP", "BO_RIMG", "Ruta Imagenes");

            DllFunciones.ProgressBar(oCompany, sboapp, 32, 1, "Creando Campo - BOPRODP - Cuenta contable compensacion ... ");
            DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 100, "", BoYesNoEnum.tNO, null, "@BOPRODP", "BO_AcctCom", "Cuen. Cont. Compen");

            DllFunciones.ProgressBar(oCompany, sboapp, 32, 1, "Creando Campo - BORTDD - Persona  ... ");
            DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 100, "", BoYesNoEnum.tNO, null, "@BORTDD", "BO_P", "Persona");

            DllFunciones.ProgressBar(oCompany, sboapp, 32, 1, "Creando Campo - BORTDD - Nombre Persona  ... ");
            DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "@BORTDD", "BO_NP", "Nombre Persona");

            DllFunciones.ProgressBar(oCompany, sboapp, 32, 1, "Creando Campo - BORTDD - Fecha Registro... ");
            DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Date, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "@BORTDD", "BO_FR", "Fecha Registro");

            DllFunciones.ProgressBar(oCompany, sboapp, 32, 1, "Creando Campo - BORTDD - Hora desde ... ");
            DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Date, BoFldSubTypes.st_Time, 254, "", BoYesNoEnum.tNO, null, "@BORTDD", "BO_TI", "Hora desde");

            DllFunciones.ProgressBar(oCompany, sboapp, 32, 1, "Creando Campo - BORTDD - Hora Hasta ... ");
            DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Date, BoFldSubTypes.st_Time, 254, "", BoYesNoEnum.tNO, null, "@BORTDD", "BO_TF", "Hora hasta");

            DllFunciones.ProgressBar(oCompany, sboapp, 32, 1, "Creando Campo - BONOPD - Codigo Articulo ... ");
            DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Date, BoFldSubTypes.st_Time, 254, "", BoYesNoEnum.tNO, null, "@BONOPD", "BO_ItemCode", "Codigo articulo");

            DllFunciones.ProgressBar(oCompany, sboapp, 32, 1, "Creando Campo - BONOPD - Descripción ... ");
            DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Date, BoFldSubTypes.st_Time, 254, "", BoYesNoEnum.tNO, null, "@BONOPD", "BO_Description", "Descripcion");

            DllFunciones.ProgressBar(oCompany, sboapp, 32, 1, "Creando Campo - BONOPD - Cantidad ... ");
            DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Date, BoFldSubTypes.st_Time, 254, "", BoYesNoEnum.tNO, null, "@BONOPD", "BO_Quantity", "Cantidad");

            DllFunciones.ProgressBar(oCompany, sboapp, 32, 1, "Creando Campo - BOTL - Tipo de documento ... ");
            string[] ValidValuesFields2 = { "OIGE", "Salida de Mercancia", "OIGN", "Entrada de Mercancia" };
            DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 10, "", BoYesNoEnum.tNO, ValidValuesFields2, "@BOTL", "BO_TD", "Tipo documento");

            DllFunciones.ProgressBar(oCompany, sboapp, 32, 1, "Creando Campo - BOTL - Codigo de Articulo ... ");
            DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 100, "", BoYesNoEnum.tNO, null, "@BOTL", "BO_ItemCode", "Codigo de Articulo");

            DllFunciones.ProgressBar(oCompany, sboapp, 32, 1, "Creando Campo - BOTL - Almacen ... ");
            DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 100, "", BoYesNoEnum.tNO, null, "@BOTL", "BO_WhsCode", "Almacen");

            DllFunciones.ProgressBar(oCompany, sboapp, 32, 1, "Creando Campo - BOTL - DocEntry Salida ... ");
            DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 100, "", BoYesNoEnum.tNO, null, "@BOTL", "BO_DocEntryOIGE", "DocEntry Salida");

            DllFunciones.ProgressBar(oCompany, sboapp, 32, 1, "Creando Campo - BOTL - DocEntry Entrada ... ");
            DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 100, "", BoYesNoEnum.tNO, null, "@BOTL", "BO_DocEntryOIGN", "DocEntry Entrada");

            DllFunciones.ProgressBar(oCompany, sboapp, 32, 1, "Creando Campo - BORPLI - Codigo Articulo ... ");
            DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 100, "", BoYesNoEnum.tNO, null, "@BORPLI", "BO_ItemCode", "Codigo Articulo");

            #endregion

            #region Creacion de UDOS 

            DllFunciones.ProgressBar(oCompany, sboapp, 32, 1, "Creando UDO - Registro de tiempo detallado..");
            string[] TablaseBilling = { "BORTDC", "BORTDD" };
            DllFunciones.CrearUDO(oCompany, sboapp, "BORTD", "BO Registro Tiempos", BoUDOObjType.boud_Document, TablaseBilling, BoYesNoEnum.tNO, BoYesNoEnum.tYES, null, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, 0, 1, BoYesNoEnum.tYES, "BO_BORTD_Log");

            DllFunciones.ProgressBar(oCompany, sboapp, 32, 1, "Creando UDO - Registro de tiempo detallado..");
            string[] TablaseBilling1 = { "BORPEN", "BORPLI" };
            DllFunciones.CrearUDO(oCompany, sboapp, "BORP", "BORP", BoUDOObjType.boud_MasterData, TablaseBilling1, BoYesNoEnum.tNO, BoYesNoEnum.tYES, null, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, 0, 1, BoYesNoEnum.tNO, null);


            #endregion

            #region Creacion de procedimientos almacenados

            DllFunciones.ProgressBar(oCompany, sboapp, 32, 1, "Creando Procedimientos almacenados , por favor espere...");
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

        public Boolean Create_Order_Prodcution(SAPbouiCOM.Application sboapp, SAPbobsCOM.Company oCompany, string _sMotor, string sUserSignature, SAPbouiCOM.Form oFormNWO)
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

                string sGetSeriesNumberProduction;
                string sGetNextDocNum;

                SAPbobsCOM.Recordset oGetSeriesNumberProduction = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset oGetNextDocNum = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

                SAPbouiCOM.Matrix oMatrixNWO = (SAPbouiCOM.Matrix)oFormNWO.Items.Item("mtxRP").Specific;
                SAPbobsCOM.ProductionOrders oWorkOrder = (SAPbobsCOM.ProductionOrders)oCompany.GetBusinessObject(BoObjectTypes.oProductionOrders);

                #endregion

                #region Obtiene serie de numeracion activas para produccion

                sGetSeriesNumberProduction = DllFunciones.GetStringXMLDocument(oCompany, "BOProduction", "Production", "GetSNProduction");
                oGetSeriesNumberProduction.DoQuery(sGetSeriesNumberProduction);

                #endregion

                #region Obtiene el consecutivo del documento a crear

                sGetNextDocNum = DllFunciones.GetStringXMLDocument(oCompany, "BOProduction", "Production", "GetNextDocNum");
                oGetNextDocNum.DoQuery(sGetNextDocNum);

                #endregion

                #region Crea la orden de produccion y sus semielaborados

                for (int i = 1; i <= oMatrixNWO.VisualRowCount; i++)
                {

                    string sTipoOrden = ((SAPbouiCOM.EditText)(oMatrixNWO.Columns.Item("Col_0").Cells.Item(i).Specific)).Value;
                    string sArticulo = ((SAPbouiCOM.EditText)(oMatrixNWO.Columns.Item("Col_1").Cells.Item(i).Specific)).Value;
                    string sQuantity = ((SAPbouiCOM.EditText)(oMatrixNWO.Columns.Item("Col_3").Cells.Item(i).Specific)).Value;
                    sQuantity = sQuantity.Replace(".", ",");
                    string sPosicion = Convert.ToString(i);
                    string sDocNum = Convert.ToString(oGetNextDocNum.Fields.Item("Consecutivo").Value.ToString());

                    if (sTipoOrden == "P. Terminado")
                    {
                        #region Crea la orden de produccion Principal

                        oWorkOrder.ItemNo = sArticulo;
                        oWorkOrder.Series = Convert.ToInt32(oGetSeriesNumberProduction.Fields.Item("SNPT").Value.ToString());
                        oWorkOrder.StartDate = DateTime.Now;
                        oWorkOrder.ProductionOrderType = BoProductionOrderTypeEnum.bopotStandard;
                        oWorkOrder.PlannedQuantity = Convert.ToDouble(sQuantity);
                        //oWorkOrder.Warehouse = Convert.ToString(oGetWorkOrderLine.Fields.Item("wareHouse").Value.ToString());
                        oWorkOrder.UserFields.Fields.Item("U_BO_TO").Value = "T";
                        oWorkOrder.UserFields.Fields.Item("U_BO_OPP").Value = sDocNum;
                        oWorkOrder.UserFields.Fields.Item("U_BO_PosId").Value = Convert.ToString(i);
                        int Rsd = oWorkOrder.Add();

                        if (Rsd == 0)
                        {
                            if (oMatrixNWO.VisualRowCount == 1)
                            {
                                DllFunciones.StatusBar(sboapp, BoStatusBarMessageType.smt_Success, "Ruta de produccion creada correctamente");
                            }
                            else
                            {
                                DllFunciones.ProgressBar(oCompany, sboapp, oMatrixNWO.VisualRowCount, 1, "Creando ruta de produción, por favor espere...");
                            }
                        }
                        else
                        {
                            DllFunciones.sendMessageBox(sboapp, oCompany.GetLastErrorDescription());
                        }

                        #endregion
                    }
                    else
                    {
                        #region Crea la orden de produccion Semielaborado

                        oWorkOrder.ItemNo = sArticulo;
                        oWorkOrder.Series = Convert.ToInt32(oGetSeriesNumberProduction.Fields.Item("SNPS").Value.ToString());
                        oWorkOrder.StartDate = DateTime.Now;
                        oWorkOrder.ProductionOrderType = BoProductionOrderTypeEnum.bopotStandard;
                        oWorkOrder.PlannedQuantity = Convert.ToDouble(sQuantity);
                        //oWorkOrder.Warehouse = Convert.ToString(oGetWorkOrderLine.Fields.Item("wareHouse").Value.ToString());
                        oWorkOrder.UserFields.Fields.Item("U_BO_TO").Value = "S";
                        oWorkOrder.UserFields.Fields.Item("U_BO_OPP").Value = sDocNum;
                        oWorkOrder.UserFields.Fields.Item("U_BO_PosId").Value = Convert.ToString(i);
                        int Rsd = oWorkOrder.Add();

                        if (Rsd == 0)
                        {
                            DllFunciones.ProgressBar(oCompany, sboapp, oMatrixNWO.VisualRowCount, 1, "Creando ruta de produción, por favor espere...");

                            if (i == oMatrixNWO.VisualRowCount)
                            {
                                DllFunciones.StatusBar(sboapp, BoStatusBarMessageType.smt_Success, "Ruta de produccion creada correctamente");
                            }
                        }
                        else
                        {
                            DllFunciones.sendMessageBox(sboapp, oCompany.GetLastErrorDescription());
                        }

                        #endregion
                    }
                }

                #endregion


                return Flag = true;
            }
            catch (Exception e)
            {

                DllFunciones.sendErrorMessage(sboapp, e);
                Flag = false;
            }
            return Flag;

        }

        public void LoadFormParProduction(SAPbouiCOM.Application sboapp, SAPbobsCOM.Company oCompany, SAPbouiCOM.Form oFormParProduction)
        {
            Funciones.Comunes DLLFunciones = new Funciones.Comunes();

            try
            {
                #region Variables y objetos

                SAPbouiCOM.ComboBox _cboSNPT = (SAPbouiCOM.ComboBox)oFormParProduction.Items.Item("txtSNPT").Specific;
                SAPbouiCOM.ComboBox _cboSNPS = (SAPbouiCOM.ComboBox)oFormParProduction.Items.Item("txtSNPS").Specific;
                SAPbouiCOM.ComboBox _cboSNGI = (SAPbouiCOM.ComboBox)oFormParProduction.Items.Item("txtSNGI").Specific;
                SAPbouiCOM.ComboBox _cboSNGR = (SAPbouiCOM.ComboBox)oFormParProduction.Items.Item("txtSNGR").Specific;
                SAPbouiCOM.EditText _txtAcct = (SAPbouiCOM.EditText)oFormParProduction.Items.Item("txtAcct").Specific;

                string sNumberSeriesActive = null;
                string sNumberSeriesSAPWorkOrder = null;
                string sNumberSeriesSAPGoodReceipt = null;
                string sNumberSeriesSAPGoodIssue = null;

                SAPbouiCOM.PictureBox oLogoBO = (SAPbouiCOM.PictureBox)oFormParProduction.Items.Item("imgLogoBO").Specific;
                SAPbobsCOM.Recordset oValidValuesSNActive = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset oValidValuesSNSAPWorkOrder = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset oValidValuesSNSAPGoodIssue = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset oValidValuesSNSAPGoodReceipt = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                #endregion

                #region Centra en pantalla formulario

                oFormParProduction.Left = (sboapp.Desktop.Width - oFormParProduction.Width) / 2;
                oFormParProduction.Top = (sboapp.Desktop.Height - oFormParProduction.Height) / 4;

                #endregion

                #region Asignacion Logo BO

                oLogoBO.Picture = (Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Core\\Imagenes\\BO.jpg");

                #endregion

                #region Adicion de DataSource

                oFormParProduction.DataSources.UserDataSources.Add("ACCTDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);

                _txtAcct.DataBind.SetBound(true, "", "ACCTDS");

                #endregion

                #region Busqueda de series de numeracion asignada

                sNumberSeriesActive = DLLFunciones.GetStringXMLDocument(oCompany, "BOProduction", "Production", "GetNumberSerieActive");

                sNumberSeriesSAPWorkOrder = DLLFunciones.GetStringXMLDocument(oCompany, "BOProduction", "Production", "GetNumberSerieSAPWorkOrder");
                sNumberSeriesSAPGoodIssue = DLLFunciones.GetStringXMLDocument(oCompany, "BOProduction", "Production", "GetNumberSerieSAPGoodsIssue");
                sNumberSeriesSAPGoodReceipt = DLLFunciones.GetStringXMLDocument(oCompany, "BOProduction", "Production", "GetNumberSerieSAPGoodsReceipt");

                oValidValuesSNActive.DoQuery(sNumberSeriesActive);
                oValidValuesSNSAPWorkOrder.DoQuery(sNumberSeriesSAPWorkOrder);
                oValidValuesSNSAPGoodIssue.DoQuery(sNumberSeriesSAPGoodIssue);
                oValidValuesSNSAPGoodReceipt.DoQuery(sNumberSeriesSAPGoodReceipt);

                #endregion

                #region Valores Series de numeración

                if (oValidValuesSNActive.RecordCount > 0)
                {
                    #region Busca las series de numeracion para Work Order

                    oValidValuesSNSAPWorkOrder.MoveFirst();

                    do
                    {
                        _cboSNPT.ValidValues.Add(oValidValuesSNSAPWorkOrder.Fields.Item("Code_SNPSAP").Value.ToString(), oValidValuesSNSAPWorkOrder.Fields.Item("Name_SNPSAP").Value.ToString());
                        _cboSNPS.ValidValues.Add(oValidValuesSNSAPWorkOrder.Fields.Item("Code_SNPSAP").Value.ToString(), oValidValuesSNSAPWorkOrder.Fields.Item("Name_SNPSAP").Value.ToString());

                        oValidValuesSNSAPWorkOrder.MoveNext();


                    } while (oValidValuesSNSAPWorkOrder.EoF == false);

                    #endregion

                    #region Busca las series de numeracion para Salidas de mercancia

                    oValidValuesSNSAPGoodIssue.MoveFirst();

                    do
                    {
                        _cboSNGI.ValidValues.Add(oValidValuesSNSAPGoodIssue.Fields.Item("Code_SNSM").Value.ToString(), oValidValuesSNSAPGoodIssue.Fields.Item("Name_SNSM").Value.ToString());

                        oValidValuesSNSAPGoodIssue.MoveNext();


                    } while (oValidValuesSNSAPGoodIssue.EoF == false);

                    #endregion

                    #region Busca las series de numeracion para Entrada de Mercancia

                    oValidValuesSNSAPGoodReceipt.MoveFirst();

                    do
                    {
                        _cboSNGR.ValidValues.Add(oValidValuesSNSAPGoodReceipt.Fields.Item("Code_SNEM").Value.ToString(), oValidValuesSNSAPGoodReceipt.Fields.Item("Name_SNEM").Value.ToString());

                        oValidValuesSNSAPGoodReceipt.MoveNext();

                    } while (oValidValuesSNSAPGoodReceipt.EoF == false);

                    #endregion

                    oValidValuesSNActive.MoveFirst();

                    _cboSNPT.Select(oValidValuesSNActive.Fields.Item("Code_SNPT").Value.ToString(), BoSearchKey.psk_ByValue);
                    _cboSNPS.Select(oValidValuesSNActive.Fields.Item("Code_SNPS").Value.ToString(), BoSearchKey.psk_ByValue);
                    _cboSNGI.Select(oValidValuesSNActive.Fields.Item("Code_SNSM").Value.ToString(), BoSearchKey.psk_ByValue);
                    _cboSNGR.Select(oValidValuesSNActive.Fields.Item("Code_SNEM").Value.ToString(), BoSearchKey.psk_ByValue);

                    oFormParProduction.DataSources.UserDataSources.Item("ACCTDS").ValueEx = Convert.ToString(oValidValuesSNActive.Fields.Item("Cuenta_Compensacion").Value.ToString());

                }
                else
                {
                    #region Busca las series de numeracion para Work Order

                    oValidValuesSNSAPWorkOrder.MoveFirst();

                    do
                    {
                        _cboSNPT.ValidValues.Add(oValidValuesSNSAPWorkOrder.Fields.Item("Code_SNPSAP").Value.ToString(), oValidValuesSNSAPWorkOrder.Fields.Item("Name_SNPSAP").Value.ToString());
                        _cboSNPS.ValidValues.Add(oValidValuesSNSAPWorkOrder.Fields.Item("Code_SNPSAP").Value.ToString(), oValidValuesSNSAPWorkOrder.Fields.Item("Name_SNPSAP").Value.ToString());

                        oValidValuesSNSAPWorkOrder.MoveNext();


                    } while (oValidValuesSNSAPWorkOrder.EoF == false);

                    #endregion

                    #region Busca las series de numeracion para Salidas de mercancia

                    oValidValuesSNSAPGoodIssue.MoveFirst();

                    do
                    {
                        _cboSNGI.ValidValues.Add(oValidValuesSNSAPGoodIssue.Fields.Item("Code_SNSM").Value.ToString(), oValidValuesSNSAPGoodIssue.Fields.Item("Name_SNSM").Value.ToString());

                        oValidValuesSNSAPGoodIssue.MoveNext();


                    } while (oValidValuesSNSAPGoodIssue.EoF == false);

                    #endregion

                    #region Busca las series de numeracion para Entrada de Mercancia

                    oValidValuesSNSAPGoodReceipt.MoveFirst();

                    do
                    {
                        _cboSNGR.ValidValues.Add(oValidValuesSNSAPGoodReceipt.Fields.Item("Code_SNEM").Value.ToString(), oValidValuesSNSAPGoodReceipt.Fields.Item("Name_SNEM").Value.ToString());

                        oValidValuesSNSAPGoodReceipt.MoveNext();

                    } while (oValidValuesSNSAPGoodReceipt.EoF == false);

                    #endregion

                }
                #endregion

                #region Se Adiciona ChooFromList

                AddChooseFromListAccount(sboapp, oFormParProduction);

                _txtAcct.ChooseFromListUID = "CFL1";
                _txtAcct.ChooseFromListAlias = "AcctCode";

                #endregion

                #region Selecciona Folder1

                SAPbouiCOM.Folder oFolder = (SAPbouiCOM.Folder)oFormParProduction.Items.Item("Folder1").Specific;
                oFolder.Item.Click();

                #endregion

                #region Liberar Objetos

                DLLFunciones.liberarObjetos(oValidValuesSNActive);
                DLLFunciones.liberarObjetos(oValidValuesSNSAPWorkOrder);
                DLLFunciones.liberarObjetos(oValidValuesSNSAPGoodIssue);
                DLLFunciones.liberarObjetos(oValidValuesSNSAPGoodReceipt);

                #endregion

                oFormParProduction.Visible = true;
                oFormParProduction.Refresh();

            }
            catch (Exception e)
            {
                DLLFunciones.sendErrorMessage(sboapp, e);

            }
        }

        public void LoadFormBatchManagement(SAPbouiCOM.Application sboapp, SAPbobsCOM.Company oCompany, SAPbouiCOM.Form oFormBatchMagnament)
        {
            Funciones.Comunes DLLFunciones = new Funciones.Comunes();

            try
            {
                #region Variables y objetos

                SAPbouiCOM.PictureBox oLogoBO = (SAPbouiCOM.PictureBox)oFormBatchMagnament.Items.Item("imgLogoBO").Specific;
                SAPbouiCOM.EditText otxtCA = (SAPbouiCOM.EditText)oFormBatchMagnament.Items.Item("txtCA").Specific;
                SAPbouiCOM.EditText otxtDesc = (SAPbouiCOM.EditText)oFormBatchMagnament.Items.Item("txtDesc").Specific;

                SAPbouiCOM.EditText otxtWH = (SAPbouiCOM.EditText)oFormBatchMagnament.Items.Item("txtWH").Specific;
                SAPbouiCOM.EditText otxtWHD = (SAPbouiCOM.EditText)oFormBatchMagnament.Items.Item("txtWHD").Specific;

                SAPbouiCOM.Matrix oMtxGL = (SAPbouiCOM.Matrix)oFormBatchMagnament.Items.Item("mtxLO").Specific;
                SAPbouiCOM.Matrix oMtxLD = (SAPbouiCOM.Matrix)oFormBatchMagnament.Items.Item("mtxLD").Specific;
                #endregion

                #region Centra en pantalla formulario

                oFormBatchMagnament.Left = (sboapp.Desktop.Width - oFormBatchMagnament.Width) / 2;
                oFormBatchMagnament.Top = (sboapp.Desktop.Height - oFormBatchMagnament.Height) / 4;

                #endregion

                #region Asignacion Logo BO

                oLogoBO.Picture = (Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Core\\Imagenes\\BO.jpg");

                #endregion

                #region Adicion de DataSource

                oFormBatchMagnament.DataSources.UserDataSources.Add("CADS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
                oFormBatchMagnament.DataSources.UserDataSources.Add("DDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);

                oFormBatchMagnament.DataSources.UserDataSources.Add("WHDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
                oFormBatchMagnament.DataSources.UserDataSources.Add("WHDDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);

                oFormBatchMagnament.DataSources.UserDataSources.Add("#1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oFormBatchMagnament.DataSources.UserDataSources.Add("DSCol0", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oFormBatchMagnament.DataSources.UserDataSources.Add("DSCol1", SAPbouiCOM.BoDataType.dt_QUANTITY, 100);
                oFormBatchMagnament.DataSources.UserDataSources.Add("DSCol2", SAPbouiCOM.BoDataType.dt_QUANTITY, 100);

                oFormBatchMagnament.DataSources.UserDataSources.Add("#2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oFormBatchMagnament.DataSources.UserDataSources.Add("DSCol4", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oFormBatchMagnament.DataSources.UserDataSources.Add("DSCol6", SAPbouiCOM.BoDataType.dt_QUANTITY, 100);

                otxtCA.DataBind.SetBound(true, "", "CADS");
                otxtDesc.DataBind.SetBound(true, "", "DDS");

                otxtWH.DataBind.SetBound(true, "", "WHDS");
                otxtWHD.DataBind.SetBound(true, "", "WHDDS");

                oMtxGL.Columns.Item("#").DataBind.SetBound(true, "", "#1");
                oMtxGL.Columns.Item("Col_0").DataBind.SetBound(true, "", "DSCol0");
                oMtxGL.Columns.Item("Col_1").DataBind.SetBound(true, "", "DSCol1");
                oMtxGL.Columns.Item("Col_2").DataBind.SetBound(true, "", "DSCol2");

                oMtxLD.Columns.Item("#").DataBind.SetBound(true, "", "#2");
                oMtxLD.Columns.Item("Col_0").DataBind.SetBound(true, "", "DSCol4");
                oMtxLD.Columns.Item("Col_2").DataBind.SetBound(true, "", "DSCol6");

                #endregion

                #region Se adicona el ChooFromList 

                AddCFLOITMBatchNumber(sboapp, oFormBatchMagnament);

                AddCFLWareHouse(sboapp, oFormBatchMagnament);

                otxtCA.ChooseFromListUID = "CFL1";
                otxtCA.ChooseFromListAlias = "ItemCode";

                otxtWH.ChooseFromListUID = "CFL3";
                otxtWH.ChooseFromListAlias = "WhsCode";

                #endregion

                #region Deshabilita items formularios

                SAPbouiCOM.Button Obtn1 = (SAPbouiCOM.Button)oFormBatchMagnament.Items.Item("btn1").Specific;
                Obtn1.Item.Enabled = false;

                SAPbouiCOM.Button Obtn2 = (SAPbouiCOM.Button)oFormBatchMagnament.Items.Item("btn2").Specific;
                Obtn2.Item.Enabled = false;

                SAPbouiCOM.Button Obtn3 = (SAPbouiCOM.Button)oFormBatchMagnament.Items.Item("btn3").Specific;
                Obtn3.Item.Enabled = false;

                SAPbouiCOM.Button Obtn4 = (SAPbouiCOM.Button)oFormBatchMagnament.Items.Item("btn4").Specific;
                Obtn4.Item.Enabled = false;

                SAPbouiCOM.Button obtnTL = (SAPbouiCOM.Button)oFormBatchMagnament.Items.Item("btnTL").Specific;
                obtnTL.Item.Enabled = false;

                #endregion

                oFormBatchMagnament.Visible = true;
                oFormBatchMagnament.Refresh();

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
                SAPbouiCOM.ComboBox ocboEOP = (SAPbouiCOM.ComboBox)oFormControlProduction.Items.Item("cboEOP").Specific;
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

                #region Asigna valores por defecto en los campos

                ocboEOP.Select("-", BoSearchKey.psk_ByValue);

                #endregion

                #region Carga Informacion al Matrix

                sConsultaOP = DLLFunciones.GetStringXMLDocument(oCompany, "BOProduction", "Production", "GetWorkOrders");

                sConsultaOP = sConsultaOP.Replace("%ItemCode%", "");

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
                            oMatrixCOP.CommonSetting.SetCellBackColor(i, 11, DLLFunciones.ColorSB1_NARANJA());
                        }

                        if (((SAPbouiCOM.EditText)oMatrixCOP.Columns.Item("Col_16").Cells.Item(i).Specific).Value == "Planificado")
                        {
                            oMatrixCOP.CommonSetting.SetCellBackColor(i, 11, DLLFunciones.ColorSB1_AGUA());
                        }

                        if (((SAPbouiCOM.EditText)oMatrixCOP.Columns.Item("Col_16").Cells.Item(i).Specific).Value == "Cerrado")
                        {
                            oMatrixCOP.CommonSetting.SetCellBackColor(i, 11, DLLFunciones.ColorSB1_LIMA());
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

                oFormControlProduction.State = BoFormStateEnum.fs_Maximized;
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

        public void FormPO(SAPbouiCOM.Form _oFormPO)
        {
            _oFormPO.Title = "Servicio Externo";
        }

        public void ChangePaneFolderWorkOrder(SAPbouiCOM.Form oFormWorkOrder)
        {
            SAPbouiCOM.Form _oFormWorkOrder;
            _oFormWorkOrder = oFormWorkOrder;
            _oFormWorkOrder.PaneLevel = 28;
        }

        public void ChangueFormProductionRoute(SAPbouiCOM.Application sboapp, SAPbobsCOM.Company oCompany, SAPbouiCOM.Form oFormProductionRoute)
        {
            Funciones.Comunes DLLFunciones = new Funciones.Comunes();

            try
            {
                #region Variables y objetos

                SAPbouiCOM.Grid oGridRP = (Grid)oFormProductionRoute.Items.Item("GridRP").Specific;
                SAPbouiCOM.PictureBox oLogoBO = (SAPbouiCOM.PictureBox)oFormProductionRoute.Items.Item("imgLogoBO").Specific;
                SAPbouiCOM.EditTextColumn oLinkedButton1;
                SAPbouiCOM.EditTextColumn oLinkedButton2;

                #endregion

                #region Centra en pantalla formulario

                oFormProductionRoute.Left = (sboapp.Desktop.Width - oFormProductionRoute.Width) / 2;
                oFormProductionRoute.Top = (sboapp.Desktop.Height - oFormProductionRoute.Height) / 4;

                #endregion

                #region Asignacion Logo BO

                oLogoBO.Picture = (Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Core\\Imagenes\\BO.jpg");

                #endregion

                #region Carga Informacion al Grid

                string sGetProductionRouteStructure = DLLFunciones.GetStringXMLDocument(oCompany, "BOProduction", "Production", "GetProductionRouteStructure");
                sGetProductionRouteStructure = sGetProductionRouteStructure.Replace(" WHERE T0.Code LIKE '%ItemCode%'", "");

                oFormProductionRoute.DataSources.DataTables.Add("DT_PR1");
                oFormProductionRoute.DataSources.DataTables.Item(0).ExecuteQuery(sGetProductionRouteStructure);
                oGridRP.DataTable = oFormProductionRoute.DataSources.DataTables.Item("DT_PR1");

                oGridRP.Columns.Item(0).Editable = false;
                oLinkedButton1 = ((SAPbouiCOM.EditTextColumn)(oGridRP.Columns.Item(0)));
                oLinkedButton1.LinkedObjectType = "4";

                oGridRP.Columns.Item(1).Editable = false;
                oLinkedButton2 = ((SAPbouiCOM.EditTextColumn)(oGridRP.Columns.Item(1)));
                oLinkedButton2.LinkedObjectType = "4";

                oGridRP.Columns.Item(2).Editable = false;
                oGridRP.Columns.Item(3).Editable = false;

                oGridRP.Columns.Item(4).Editable = false;
                oGridRP.Columns.Item(4).RightJustified = true;

                oGridRP.Columns.Item(5).Editable = false;
                oGridRP.Columns.Item(5).RightJustified = true;

                oGridRP.CollapseLevel = 2;

                oGridRP.Rows.CollapseAll();

                #endregion

                oFormProductionRoute.State = BoFormStateEnum.fs_Maximized;
                oFormProductionRoute.Visible = true;
                oFormProductionRoute.Refresh();

                DLLFunciones.liberarObjetos(oGridRP);

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


                if (string.IsNullOrEmpty(sDocEntryOPS))
                {
                    sDocEntryOPS = ((SAPbouiCOM.EditText)oMatrixCOP.Columns.Item("Col_20").Cells.Item(pVal.Row).Specific).Value;
                }


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

        public void LoadFormNewWorkOrder(SAPbouiCOM.Application sboapp, SAPbobsCOM.Company oCompany, SAPbouiCOM.Form oFormNewWorkOrder, SAPbouiCOM.ItemEvent pVal)
        {
            #region Variables y objetos

            SAPbouiCOM.Matrix oMatrixNOP = (SAPbouiCOM.Matrix)oFormNewWorkOrder.Items.Item("mtxRP").Specific;

            SAPbouiCOM.DataTable oTableWO = oFormNewWorkOrder.DataSources.DataTables.Item("DT_WO");

            SAPbouiCOM.PictureBox oLogoBO = (SAPbouiCOM.PictureBox)oFormNewWorkOrder.Items.Item("imgLogoBO").Specific;

            SAPbouiCOM.EditText otxtIPR = (SAPbouiCOM.EditText)oFormNewWorkOrder.Items.Item("txtIPR").Specific;

            #endregion

            #region Centra en pantalla formulario

            oFormNewWorkOrder.Left = (sboapp.Desktop.Width - oFormNewWorkOrder.Width) / 2;
            oFormNewWorkOrder.Top = (sboapp.Desktop.Height - oFormNewWorkOrder.Height) / 4;

            #endregion

            #region Asignacion Logo BO

            oLogoBO.Picture = (Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Core\\Imagenes\\BO.jpg");

            #endregion

            #region Adicion de DataSource

            oFormNewWorkOrder.DataSources.DataTables.Add("DTRP");

            oFormNewWorkOrder.DataSources.UserDataSources.Add("IRPDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);

            oFormNewWorkOrder.DataSources.UserDataSources.Add("#", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
            oFormNewWorkOrder.DataSources.UserDataSources.Add("DSCol0", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
            oFormNewWorkOrder.DataSources.UserDataSources.Add("DSCol1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
            oFormNewWorkOrder.DataSources.UserDataSources.Add("DSCol2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
            oFormNewWorkOrder.DataSources.UserDataSources.Add("DSCol3", SAPbouiCOM.BoDataType.dt_QUANTITY, 100);

            otxtIPR.DataBind.SetBound(true, "", "IRPDS");

            oMatrixNOP.Columns.Item("#").DataBind.SetBound(true, "", "#");
            oMatrixNOP.Columns.Item("Col_0").DataBind.SetBound(true, "", "DSCol0");
            oMatrixNOP.Columns.Item("Col_1").DataBind.SetBound(true, "", "DSCol1");
            oMatrixNOP.Columns.Item("Col_2").DataBind.SetBound(true, "", "DSCol2");
            oMatrixNOP.Columns.Item("Col_3").DataBind.SetBound(true, "", "DSCol3");

            #endregion

            #region Se adicona el ChooFromList 

            AddCFLProductionRouteNWO(sboapp, oFormNewWorkOrder);
            AddChooseFromListoOITMMatrix(sboapp, oFormNewWorkOrder);

            otxtIPR.ChooseFromListUID = "CFL1";
            otxtIPR.ChooseFromListAlias = "Code";

            oMatrixNOP.Columns.Item("Col_1").ChooseFromListUID = "CFL3";
            oMatrixNOP.Columns.Item("Col_1").ChooseFromListAlias = "CardCode";

            #endregion         

            #region Adicionar primera linea en la Matrix

            oMatrixNOP.AddRow();

            oMatrixNOP.Columns.Item("Col_1").Cells.Item(1).Click();

            #endregion

            oFormNewWorkOrder.Refresh();
            oFormNewWorkOrder.Visible = true;

        }

        public void LoadFormNewProductionRoute(SAPbouiCOM.Application sboapp, SAPbobsCOM.Company oCompany, SAPbouiCOM.Form oFormNewProductionRoute, SAPbouiCOM.ItemEvent pVal)
        {

            #region Variables y objetos

            SAPbouiCOM.PictureBox oLogoBO = (SAPbouiCOM.PictureBox)oFormNewProductionRoute.Items.Item("imgLogoBO").Specific;

            SAPbouiCOM.EditText otxtItemCode = (SAPbouiCOM.EditText)oFormNewProductionRoute.Items.Item("0_U_E").Specific;

            SAPbouiCOM.Matrix oMatrixRP = (SAPbouiCOM.Matrix)oFormNewProductionRoute.Items.Item("0_U_G").Specific;

            #endregion

            #region Centra en pantalla formulario

            oFormNewProductionRoute.Left = (sboapp.Desktop.Width - oFormNewProductionRoute.Width) / 2;
            oFormNewProductionRoute.Top = (sboapp.Desktop.Height - oFormNewProductionRoute.Height) / 4;

            #endregion

            #region Asignacion Logo BO

            oLogoBO.Picture = (Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Core\\Imagenes\\BO.jpg");

            #endregion

            #region Se adicona el ChooFromList 

            AddChooseFromListoOITM(sboapp, oFormNewProductionRoute);
            AddChooseFromListoOITMMatrix(sboapp, oFormNewProductionRoute);

            otxtItemCode.ChooseFromListUID = "CFL1";
            otxtItemCode.ChooseFromListAlias = "ItemCode";

            oMatrixRP.Columns.Item("C_0_2").ChooseFromListUID = "CFL3";
            oMatrixRP.Columns.Item("C_0_2").ChooseFromListAlias = "ItemCode";

            oFormNewProductionRoute.Refresh();

            #endregion

            #region Se adiciona Primera linea

            oMatrixRP.AddRow();

            #endregion

        }

        public void CFLAfterMatrix(string _FormUID, SAPbouiCOM.Form _FormWO, ItemEvent pVal, SAPbobsCOM.Company _oCompany, SAPbouiCOM.Application _sboapp)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            #region Variables y Objetos

            string sCFL_ID = null;
            string _sMotorDB = (Convert.ToString(_oCompany.DbServerType));
            int NumeroLinea;

            SAPbouiCOM.CellPosition _Cell;

            SAPbouiCOM.Matrix oMatrixWO = (SAPbouiCOM.Matrix)_FormWO.Items.Item("mtxRP").Specific;

            _Cell = oMatrixWO.GetCellFocus();

            NumeroLinea = _Cell.rowIndex;

            SAPbouiCOM.IChooseFromListEvent oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));

            sCFL_ID = oCFLEvento.ChooseFromListUID;

            SAPbouiCOM.Form _oFormWO = _sboapp.Forms.Item(_FormUID);

            SAPbouiCOM.ChooseFromList oCFL = _oFormWO.ChooseFromLists.Item(sCFL_ID);

            #endregion

            if (oCFLEvento.BeforeAction == false)
            {

                #region Variables y Objetos 

                string Col_0 = null;
                string Col_1 = null;
                string Col_2 = null;


                SAPbouiCOM.DataTable oDataTable = oCFLEvento.SelectedObjects;

                #endregion

                try
                {

                    if (oDataTable == null)
                    {
                    }
                    else
                    {
                        Col_1 = System.Convert.ToString(oDataTable.GetValue(0, 0));
                        Col_2 = System.Convert.ToString(oDataTable.GetValue(1, 0));

                        if (_sMotorDB == "dst_MSSQL2012" || _sMotorDB == "dst_MSSQL2014" || _sMotorDB == "dst_MSSQL2016" || _sMotorDB == "dst_MSSQL2017" || _sMotorDB == "dst_HANADB")
                        {
                            if (pVal.ItemUID == "mtxRP" & pVal.ColUID == "Col_1")
                            {
                                if (pVal.Row == 1)
                                {
                                    Col_0 = "P. Terminado";

                                    _FormWO.DataSources.UserDataSources.Item("DSCol0").ValueEx = Col_0;
                                }
                                else
                                {
                                    Col_0 = "P. Semielaborado";

                                    _FormWO.DataSources.UserDataSources.Item("DSCol0").ValueEx = Col_0;

                                }

                                _FormWO.DataSources.UserDataSources.Item("#").ValueEx = System.Convert.ToString(pVal.Row);
                                _FormWO.DataSources.UserDataSources.Item("DSCol1").ValueEx = Col_1;
                                _FormWO.DataSources.UserDataSources.Item("DSCol2").ValueEx = Col_2;

                                oMatrixWO.SetLineData(pVal.Row);
                            }
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

        public void CFLAfterOITMMatrix(string _FormUID, SAPbouiCOM.Form _FormWO, ItemEvent pVal, SAPbobsCOM.Company _oCompany, SAPbouiCOM.Application _sboapp)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            #region Variables y Objetos

            string sCFL_ID = null;
            string _sMotorDB = (Convert.ToString(_oCompany.DbServerType));
            int NumeroLinea;

            SAPbouiCOM.CellPosition _Cell;

            SAPbouiCOM.Matrix oMatrixRP = (SAPbouiCOM.Matrix)_FormWO.Items.Item("0_U_G").Specific;

            _Cell = oMatrixRP.GetCellFocus();

            NumeroLinea = _Cell.rowIndex;

            SAPbouiCOM.IChooseFromListEvent oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));

            sCFL_ID = oCFLEvento.ChooseFromListUID;

            SAPbouiCOM.Form _oFormWO = _sboapp.Forms.Item(_FormUID);

            SAPbouiCOM.ChooseFromList oCFL = _oFormWO.ChooseFromLists.Item(sCFL_ID);

            #endregion

            if (oCFLEvento.BeforeAction == false)
            {

                #region Variables y Objetos 

                string Col_1 = null;

                SAPbouiCOM.DataTable oDataTable = oCFLEvento.SelectedObjects;

                #endregion

                try
                {

                    if (oDataTable == null)
                    {
                    }
                    else
                    {
                        Col_1 = System.Convert.ToString(oDataTable.GetValue(0, 0));

                        if (_sMotorDB == "dst_MSSQL2012" || _sMotorDB == "dst_MSSQL2014" || _sMotorDB == "dst_MSSQL2016" || _sMotorDB == "dst_MSSQL2017" || _sMotorDB == "dst_HANADB")
                        {
                            if (pVal.ItemUID == "0_U_G" & pVal.ColUID == "C_0_2")
                            {
                                ((SAPbouiCOM.EditText)(oMatrixRP.Columns.Item("C_0_2").Cells.Item(pVal.Row).Specific)).Value = Col_1;

                            }
                        }

                    }


                }
                catch (Exception e)
                {
                    DllFunciones.sendErrorMessage(_sboapp, e);
                }

                if ((_FormUID == "CFL3") & (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD))
                {
                    System.Windows.Forms.Application.Exit();
                }
            }
        }

        public void CFLAfterIRP(string _FormUID, SAPbouiCOM.Form _FormWO, ItemEvent pVal, SAPbobsCOM.Company _oCompany, SAPbouiCOM.Application _sboapp)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            #region Variables y Objetos

            string sCFL_ID = null;
            string _sMotorDB = (Convert.ToString(_oCompany.DbServerType));

            SAPbouiCOM.EditText otxtIPR = (SAPbouiCOM.EditText)_FormWO.Items.Item("txtIPR").Specific;

            SAPbouiCOM.IChooseFromListEvent oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));

            sCFL_ID = oCFLEvento.ChooseFromListUID;

            SAPbouiCOM.Form _oFormWO = _sboapp.Forms.Item(_FormUID);

            SAPbouiCOM.ChooseFromList oCFL = _oFormWO.ChooseFromLists.Item(sCFL_ID);

            #endregion

            if (oCFLEvento.BeforeAction == false)
            {
                #region Variables y Objetos 

                string val = null;

                SAPbouiCOM.DataTable oDataTable = oCFLEvento.SelectedObjects;

                #endregion

                try
                {

                    if (oDataTable == null)
                    {
                    }
                    else
                    {
                        val = System.Convert.ToString(oDataTable.GetValue(0, 0));

                        SAPbobsCOM.Recordset oPR = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                        if (_sMotorDB == "dst_MSSQL2012" || _sMotorDB == "dst_MSSQL2014" || _sMotorDB == "dst_MSSQL2016" || _sMotorDB == "dst_MSSQL2017" || _sMotorDB == "dst_HANADB")
                        {
                            _oFormWO.DataSources.UserDataSources.Item("IRPDS").ValueEx = val;

                        }

                    }
                }
                catch (Exception e)
                {
                    DllFunciones.sendErrorMessage(_sboapp, e);
                }

                if ((_FormUID == "CFL3") & (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD))
                {
                    System.Windows.Forms.Application.Exit();
                }
            }
        }

        public void CFLAfterItemNewProductionRoute(string _FormUID, SAPbouiCOM.Form _FormNPR, ItemEvent pVal, SAPbobsCOM.Company _oCompany, SAPbouiCOM.Application _sboapp)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            #region Variables y Objetos

            string sCFL_ID = null;
            string _sMotorDB = (Convert.ToString(_oCompany.DbServerType));

            SAPbouiCOM.EditText otxtItemCode = (SAPbouiCOM.EditText)_FormNPR.Items.Item("0_U_E").Specific;

            SAPbouiCOM.IChooseFromListEvent oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));

            sCFL_ID = oCFLEvento.ChooseFromListUID;

            SAPbouiCOM.Form _oFormWO = _sboapp.Forms.Item(_FormUID);

            SAPbouiCOM.ChooseFromList oCFL = _oFormWO.ChooseFromLists.Item(sCFL_ID);

            #endregion

            if (oCFLEvento.BeforeAction == false)
            {
                #region Variables y Objetos 

                string val = null;

                SAPbouiCOM.DataTable oDataTable = oCFLEvento.SelectedObjects;

                #endregion

                try
                {

                    if (oDataTable == null)
                    {
                    }
                    else
                    {
                        val = System.Convert.ToString(oDataTable.GetValue(0, 0));

                        if (_sMotorDB == "dst_MSSQL2012" || _sMotorDB == "dst_MSSQL2014" || _sMotorDB == "dst_MSSQL2016" || _sMotorDB == "dst_MSSQL2017" || _sMotorDB == "dst_HANADB")
                        {
                            otxtItemCode.Value = val;
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

        public void CFLAfterOACT(string _FormUID, SAPbouiCOM.Form _FormWO, ItemEvent pVal, SAPbobsCOM.Company _oCompany, SAPbouiCOM.Application _sboapp)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            #region Variables y Objetos

            string sCFL_ID = null;
            string _sMotorDB = (Convert.ToString(_oCompany.DbServerType));


            SAPbouiCOM.EditText otxtAcct = (SAPbouiCOM.EditText)_FormWO.Items.Item("txtAcct").Specific;

            SAPbouiCOM.IChooseFromListEvent oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));

            sCFL_ID = oCFLEvento.ChooseFromListUID;

            SAPbouiCOM.Form _oFormWO = _sboapp.Forms.Item(_FormUID);

            SAPbouiCOM.ChooseFromList oCFL = _oFormWO.ChooseFromLists.Item(sCFL_ID);

            #endregion

            if (oCFLEvento.BeforeAction == false)
            {
                #region Variables y Objetos 

                string val = null;

                SAPbouiCOM.DataTable oDataTable = oCFLEvento.SelectedObjects;

                #endregion

                try
                {

                    if (oDataTable == null)
                    {
                    }
                    else
                    {
                        val = System.Convert.ToString(oDataTable.GetValue(0, 0));

                        SAPbobsCOM.Recordset oPR = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                        if (_sMotorDB == "dst_MSSQL2012" || _sMotorDB == "dst_MSSQL2014" || _sMotorDB == "dst_MSSQL2016" || _sMotorDB == "dst_MSSQL2017" || _sMotorDB == "dst_HANADB")
                        {
                            _oFormWO.DataSources.UserDataSources.Item("ACCTDS").ValueEx = val;

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

        public void CFLAfterOITMBatchMagnagemet(string _FormUID, SAPbouiCOM.Form _FormWO, ItemEvent pVal, SAPbobsCOM.Company _oCompany, SAPbouiCOM.Application _sboapp)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            #region Variables y Objetos

            string sCFL_ID = null;
            string _sMotorDB = (Convert.ToString(_oCompany.DbServerType));


            SAPbouiCOM.EditText otxtIPR = (SAPbouiCOM.EditText)_FormWO.Items.Item("txtCA").Specific;

            SAPbouiCOM.IChooseFromListEvent oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));

            sCFL_ID = oCFLEvento.ChooseFromListUID;

            SAPbouiCOM.Form _oFormWO = _sboapp.Forms.Item(_FormUID);

            SAPbouiCOM.ChooseFromList oCFL = _oFormWO.ChooseFromLists.Item(sCFL_ID);

            #endregion

            if (oCFLEvento.BeforeAction == false)
            {
                #region Variables y Objetos 

                string val = null;
                string val2 = null;

                SAPbouiCOM.DataTable oDataTable = oCFLEvento.SelectedObjects;

                #endregion

                try
                {

                    if (oDataTable == null)
                    {
                    }
                    else
                    {
                        val = System.Convert.ToString(oDataTable.GetValue(0, 0));
                        val2 = System.Convert.ToString(oDataTable.GetValue(1, 0));

                        SAPbobsCOM.Recordset oPR = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                        if (_sMotorDB == "dst_MSSQL2012" || _sMotorDB == "dst_MSSQL2014" || _sMotorDB == "dst_MSSQL2016" || _sMotorDB == "dst_MSSQL2017" || _sMotorDB == "dst_HANADB")
                        {
                            _oFormWO.DataSources.UserDataSources.Item("CADS").ValueEx = val;
                            _oFormWO.DataSources.UserDataSources.Item("DDS").ValueEx = val2;
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

        public void CFLAfterWareHouse(string _FormUID, SAPbouiCOM.Form _FormWO, ItemEvent pVal, SAPbobsCOM.Company _oCompany, SAPbouiCOM.Application _sboapp)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            #region Variables y Objetos

            string sCFL_ID = null;
            string _sMotorDB = (Convert.ToString(_oCompany.DbServerType));


            SAPbouiCOM.EditText otxtIPR = (SAPbouiCOM.EditText)_FormWO.Items.Item("txtWH").Specific;

            SAPbouiCOM.IChooseFromListEvent oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));

            sCFL_ID = oCFLEvento.ChooseFromListUID;

            SAPbouiCOM.Form _oFormWO = _sboapp.Forms.Item(_FormUID);

            SAPbouiCOM.ChooseFromList oCFL = _oFormWO.ChooseFromLists.Item(sCFL_ID);

            #endregion

            if (oCFLEvento.BeforeAction == false)
            {
                #region Variables y Objetos 

                string val = null;
                string val2 = null;

                SAPbouiCOM.DataTable oDataTable = oCFLEvento.SelectedObjects;

                #endregion

                try
                {

                    if (oDataTable == null)
                    {
                    }
                    else
                    {
                        val = System.Convert.ToString(oDataTable.GetValue(0, 0));
                        val2 = System.Convert.ToString(oDataTable.GetValue(1, 0));

                        SAPbobsCOM.Recordset oPR = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                        if (_sMotorDB == "dst_MSSQL2012" || _sMotorDB == "dst_MSSQL2014" || _sMotorDB == "dst_MSSQL2016" || _sMotorDB == "dst_MSSQL2017" || _sMotorDB == "dst_HANADB")
                        {
                            _oFormWO.DataSources.UserDataSources.Item("WHDS").ValueEx = val;
                            _oFormWO.DataSources.UserDataSources.Item("WHDDS").ValueEx = val2;
                        }

                    }
                }
                catch (Exception e)
                {
                    DllFunciones.sendErrorMessage(_sboapp, e);
                }

                if ((_FormUID == "CFL3") & (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD))
                {
                    System.Windows.Forms.Application.Exit();
                }
            }
        }

        public void AddItemsNWOMatrix(SAPbouiCOM.Form _FormWO, SAPbobsCOM.Company _oCompany, SAPbouiCOM.Application _sboapp)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            #region Variables y Objetos

            SAPbouiCOM.Matrix oMatrixNWO = (SAPbouiCOM.Matrix)_FormWO.Items.Item("mtxRP").Specific;
            SAPbouiCOM.EditText otxtIPR = (SAPbouiCOM.EditText)_FormWO.Items.Item("txtIPR").Specific;
            SAPbouiCOM.EditText otxtQTY = (SAPbouiCOM.EditText)_FormWO.Items.Item("txtQTY").Specific;

            SAPbouiCOM.DataTable oTablePR = _FormWO.DataSources.DataTables.Item("DTRP");

            string stxtIPR = null;


            #endregion

            #region Consulta Ruta de produccion 


            stxtIPR = otxtIPR.Value.ToString();

            string sIRP = DllFunciones.GetStringXMLDocument(_oCompany, "BOProduction", "Production", "GetProductionRoute");
            sIRP = sIRP.Replace("%ItemCode%", stxtIPR).Replace("%QTY%", otxtQTY.Value.ToString());

            //oRIPR.DoQuery(sIRP);

            oTablePR.ExecuteQuery(sIRP);

            #endregion

            #region  Carga informacion a la Matrix

            if (oTablePR.Rows.Count > 0)
            {
                oMatrixNWO.Clear();

                oMatrixNWO.Columns.Item("#").DataBind.Bind("DTRP", "#");
                oMatrixNWO.Columns.Item("Col_0").DataBind.Bind("DTRP", "Posicion");
                oMatrixNWO.Columns.Item("Col_1").DataBind.Bind("DTRP", "ItemCode");
                oMatrixNWO.Columns.Item("Col_2").DataBind.Bind("DTRP", "ItemName");
                oMatrixNWO.Columns.Item("Col_3").DataBind.Bind("DTRP", "Quantity");

                oMatrixNWO.LoadFromDataSource();

                oMatrixNWO.AutoResizeColumns();

            }

            #endregion

            #region Liberar Objetos 

            DllFunciones.liberarObjetos(oTablePR);

            #endregion

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

                oColumDocEntry = oMatrixCOP.Columns.Item("Col_8");

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

        public void UpdateParametersProduction(SAPbouiCOM.Application sboapp, SAPbobsCOM.Company _oCompany, SAPbouiCOM.Form _oFormParProduction)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            try
            {
                #region Variables y objetos

                SAPbobsCOM.UserTable oUDT = (SAPbobsCOM.UserTable)(_oCompany.UserTables.Item("BOPRODP"));

                SAPbouiCOM.ComboBox cboSNPT = (SAPbouiCOM.ComboBox)_oFormParProduction.Items.Item("txtSNPT").Specific;
                SAPbouiCOM.ComboBox cboSNPS = (SAPbouiCOM.ComboBox)_oFormParProduction.Items.Item("txtSNPS").Specific;
                SAPbouiCOM.ComboBox _cboSNGI = (SAPbouiCOM.ComboBox)_oFormParProduction.Items.Item("txtSNGI").Specific;
                SAPbouiCOM.ComboBox _cboSNGR = (SAPbouiCOM.ComboBox)_oFormParProduction.Items.Item("txtSNGR").Specific;
                SAPbouiCOM.EditText _txtAcct = (SAPbouiCOM.EditText)_oFormParProduction.Items.Item("txtAcct").Specific;

                SAPbouiCOM.Button btnOK = (SAPbouiCOM.Button)_oFormParProduction.Items.Item("btnUpdate").Specific;

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
                    oUDT.UserFields.Fields.Item("U_BO_SNSM").Value = _cboSNGI.Selected.Value;
                    oUDT.UserFields.Fields.Item("U_BO_SNEM").Value = _cboSNGR.Selected.Value;
                    oUDT.UserFields.Fields.Item("U_BO_AcctCom").Value = _txtAcct.Value.ToString();

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
                    oUDT.UserFields.Fields.Item("U_BO_SNSM").Value = _cboSNGI.Selected.Value;
                    oUDT.UserFields.Fields.Item("U_BO_SNEM").Value = _cboSNGR.Selected.Value;
                    oUDT.UserFields.Fields.Item("U_BO_AcctCom").Value = _txtAcct.Value.ToString();

                    oUDT.Add();

                    #endregion

                    DllFunciones.liberarObjetos(oSearchNextCode);
                }

                DllFunciones.StatusBar(sboapp, BoStatusBarMessageType.smt_Success, "Actualizado correctamente...");
                btnOK.Caption = "OK";
                _oFormParProduction.Mode = BoFormMode.fm_OK_MODE;
                _oFormParProduction.Refresh();

            }
            catch (Exception ex)
            {
                DllFunciones.sendErrorMessage(sboapp, ex);
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

                sConsultaOP = sConsultaOP.Replace("%ItemCode%", "");

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

        public void SearchBacht(SAPbouiCOM.Application sboapp, SAPbobsCOM.Company oCompany, SAPbouiCOM.Form oFormGL, ItemEvent pVal)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            try
            {
                #region Variables y objetos

                SAPbouiCOM.Button Obtn1 = (SAPbouiCOM.Button)oFormGL.Items.Item("btn1").Specific;
                SAPbouiCOM.Button Obtn2 = (SAPbouiCOM.Button)oFormGL.Items.Item("btn2").Specific;
                SAPbouiCOM.Button Obtn3 = (SAPbouiCOM.Button)oFormGL.Items.Item("btn3").Specific;
                SAPbouiCOM.Button Obtn4 = (SAPbouiCOM.Button)oFormGL.Items.Item("btn4").Specific;
                SAPbouiCOM.Button obtnTL = (SAPbouiCOM.Button)oFormGL.Items.Item("btnTL").Specific;
                SAPbouiCOM.Button obtnCons = (SAPbouiCOM.Button)oFormGL.Items.Item("btnCons").Specific;


                SAPbouiCOM.Matrix oMatrixGL = (Matrix)oFormGL.Items.Item("mtxLO").Specific;
                SAPbouiCOM.Matrix oMatrixLD = (Matrix)oFormGL.Items.Item("mtxLD").Specific;
                SAPbouiCOM.CommonSetting CS = oMatrixGL.CommonSetting;

                #endregion

                #region Obtiene codigo y lote

                SAPbouiCOM.EditText otxtCA = (SAPbouiCOM.EditText)oFormGL.Items.Item("txtCA").Specific;
                SAPbouiCOM.EditText otxtWH = (SAPbouiCOM.EditText)oFormGL.Items.Item("txtWH").Specific;

                string stxtCA = null;
                string stxtWH = null;

                stxtCA = otxtCA.Value.ToString();
                stxtWH = otxtWH.Value.ToString();

                #endregion

                #region Carga Lotes a la columna de la Matrix Lotes Origen

                #region Consulta Lotes en SAP

                string sBachNumber = null;

                SAPbobsCOM.Recordset oRBachNumber = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                sBachNumber = DllFunciones.GetStringXMLDocument(oCompany, "BOProduction", "Production", "GetBatchNumber");

                sBachNumber = sBachNumber.Replace("%ItemCode%", stxtCA).Replace("%WhsCode%", stxtWH);

                oRBachNumber.DoQuery(sBachNumber);

                #endregion

                if (oRBachNumber.RecordCount > 0)
                {
                    oMatrixGL.Clear();

                    oRBachNumber.MoveFirst();

                    do
                    {
                        oMatrixGL.Columns.Item("Col_0").ValidValues.Add(oRBachNumber.Fields.Item("DistNumber").Value.ToString(), "");
                        oRBachNumber.MoveNext();

                    } while (oRBachNumber.EoF == false);

                    oMatrixGL.AddRow();

                    oMatrixGL.SetCellFocus(1, 1);

                    #region Habilita items formularios

                    Obtn1.Item.Enabled = true;
                    Obtn2.Item.Enabled = true;
                    Obtn3.Item.Enabled = true;
                    Obtn4.Item.Enabled = true;
                    obtnTL.Item.Enabled = true;

                    #endregion

                    #region Deshabilita items formulario

                    obtnCons.Item.Enabled = false;

                    otxtCA.Item.Enabled = false;
                    otxtWH.Item.Enabled = false;

                    #endregion

                    #region Adiciona linea en matrix lotes destino

                    AddLineMatrixLD(oFormGL, pVal);

                    #endregion

                }
                else
                {
                    DllFunciones.sendMessageBox(sboapp, "No se encontraron lotes disponibles para el articulo" + otxtCA.Value.ToString());
                }

                #endregion

                oFormGL.Refresh();

                DllFunciones.liberarObjetos(oMatrixGL); ;
                DllFunciones.liberarObjetos(CS);

            }
            catch (Exception e)
            {
                DllFunciones.sendErrorMessage(sboapp, e);

            }
        }

        public void SearchProductionRoute(SAPbouiCOM.Application sboapp, SAPbobsCOM.Company oCompany, SAPbouiCOM.Form oFormProductionRoute)
        {
            Funciones.Comunes DLLFunciones = new Funciones.Comunes();

            try
            {
                #region Variables y objetos

                SAPbouiCOM.Grid oGridRP = (Grid)oFormProductionRoute.Items.Item("GridRP").Specific;
                SAPbouiCOM.EditTextColumn oLinkedButton1;
                SAPbouiCOM.EditTextColumn oLinkedButton2;

                SAPbouiCOM.EditText otxtCRP = (SAPbouiCOM.EditText)oFormProductionRoute.Items.Item("txtCRP").Specific;

                #endregion

                oFormProductionRoute.Freeze(true);

                #region Carga Informacion al Grid

                string sGetProductionRouteStructure = DLLFunciones.GetStringXMLDocument(oCompany, "BOProduction", "Production", "GetProductionRouteStructure");
                sGetProductionRouteStructure = sGetProductionRouteStructure.Replace("%ItemCode%", "%" + otxtCRP.Value.ToString() + "%");

                oGridRP.DataTable.Clear();

                oFormProductionRoute.DataSources.DataTables.Item(0).ExecuteQuery(sGetProductionRouteStructure);
                oGridRP.DataTable = oFormProductionRoute.DataSources.DataTables.Item("DT_PR1");

                oGridRP.Columns.Item(0).Editable = false;
                oLinkedButton1 = ((SAPbouiCOM.EditTextColumn)(oGridRP.Columns.Item(0)));
                oLinkedButton1.LinkedObjectType = "4";

                oGridRP.Columns.Item(1).Editable = false;
                oLinkedButton2 = ((SAPbouiCOM.EditTextColumn)(oGridRP.Columns.Item(1)));
                oLinkedButton2.LinkedObjectType = "4";

                oGridRP.Columns.Item(2).Editable = false;
                oGridRP.Columns.Item(3).Editable = false;

                oGridRP.Columns.Item(4).Editable = false;
                oGridRP.Columns.Item(4).RightJustified = true;

                oGridRP.Columns.Item(5).Editable = false;
                oGridRP.Columns.Item(5).RightJustified = true;

                oGridRP.CollapseLevel = 2;

                oGridRP.Rows.CollapseAll();

                #endregion

                oFormProductionRoute.Refresh();

                DLLFunciones.liberarObjetos(oGridRP);

                oFormProductionRoute.Freeze(false);


            }
            catch (Exception e)
            {
                DLLFunciones.sendErrorMessage(sboapp, e);

            }
        }

        public void SearchWorkOrder(SAPbouiCOM.Application sboapp, SAPbobsCOM.Company oCompany, SAPbouiCOM.Form oFormControlProduction)
        {
            Funciones.Comunes DLLFunciones = new Funciones.Comunes();

            try
            {
                #region Variables y objetos

                SAPbouiCOM.Matrix oMatrixCOP = (Matrix)oFormControlProduction.Items.Item("MtxCOP").Specific;
                SAPbouiCOM.CommonSetting CS = oMatrixCOP.CommonSetting;
                SAPbouiCOM.ComboBox ocboEOP = (SAPbouiCOM.ComboBox)oFormControlProduction.Items.Item("cboEOP").Specific;
                SAPbouiCOM.EditText otxtOP = (SAPbouiCOM.EditText)oFormControlProduction.Items.Item("txtOP#").Specific;
                SAPbouiCOM.DataTable oTableDT = oFormControlProduction.DataSources.DataTables.Item("DT_COP");

                string sConsultaOP;
                int iCount;

                #endregion

                oFormControlProduction.Freeze(true);

                #region Carga Informacion al Matrix

                sConsultaOP = DLLFunciones.GetStringXMLDocument(oCompany, "BOProduction", "Production", "GetWorkOrders");

                if (string.IsNullOrEmpty(otxtOP.Value.ToString()))
                {

                    if (oTableDT.IsEmpty == true)
                    {
                        sConsultaOP = sConsultaOP.Replace("%ItemCode%", otxtOP.Value.ToString());

                        oFormControlProduction.DataSources.DataTables.Item(0).ExecuteQuery(sConsultaOP);

                        oMatrixCOP.Clear();

                        iCount = oTableDT.Rows.Count;

                        if (oTableDT.IsEmpty == false)
                        {
                            #region Datasource


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

                            #endregion

                            for (int i = 1; i <= iCount; i++)
                            {

                                #region Pinta las columnas con colores y carga imagenes

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
                                    oMatrixCOP.CommonSetting.SetCellBackColor(i, 11, DLLFunciones.ColorSB1_NARANJA());
                                }

                                if (((SAPbouiCOM.EditText)oMatrixCOP.Columns.Item("Col_16").Cells.Item(i).Specific).Value == "Planificado")
                                {
                                    oMatrixCOP.CommonSetting.SetCellBackColor(i, 11, DLLFunciones.ColorSB1_AGUA());
                                }

                                if (((SAPbouiCOM.EditText)oMatrixCOP.Columns.Item("Col_16").Cells.Item(i).Specific).Value == "Cerrado")
                                {
                                    oMatrixCOP.CommonSetting.SetCellBackColor(i, 11, DLLFunciones.ColorSB1_LIMA());
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

                                #endregion
                            }

                            oMatrixCOP.AutoResizeColumns();
                        }
                    }
                    else
                    {

                    }

                }
                else
                {
                    sConsultaOP = sConsultaOP.Replace("%ItemCode%", otxtOP.Value.ToString());

                    oFormControlProduction.DataSources.DataTables.Item(0).ExecuteQuery(sConsultaOP);

                    oMatrixCOP.Clear();

                    iCount = oTableDT.Rows.Count;

                    if (oTableDT.IsEmpty == false)
                    {
                        #region Datasource


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

                        #endregion

                        for (int i = 1; i <= iCount; i++)
                        {

                            #region Pinta las columnas con colores y carga imagenes

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
                                oMatrixCOP.CommonSetting.SetCellBackColor(i, 11, DLLFunciones.ColorSB1_NARANJA());
                            }

                            if (((SAPbouiCOM.EditText)oMatrixCOP.Columns.Item("Col_16").Cells.Item(i).Specific).Value == "Planificado")
                            {
                                oMatrixCOP.CommonSetting.SetCellBackColor(i, 11, DLLFunciones.ColorSB1_AGUA());
                            }

                            if (((SAPbouiCOM.EditText)oMatrixCOP.Columns.Item("Col_16").Cells.Item(i).Specific).Value == "Cerrado")
                            {
                                oMatrixCOP.CommonSetting.SetCellBackColor(i, 11, DLLFunciones.ColorSB1_LIMA());
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

                            #endregion
                        }

                        oMatrixCOP.AutoResizeColumns();
                    }
                }



                #endregion

                oFormControlProduction.Refresh();

                oFormControlProduction.Freeze(false);

                oFormControlProduction.Refresh();

                DLLFunciones.liberarObjetos(oMatrixCOP);
                DLLFunciones.liberarObjetos(oTableDT);
                DLLFunciones.liberarObjetos(CS);

            }
            catch (Exception e)
            {
                DLLFunciones.sendErrorMessage(sboapp, e);

            }
        }

        public void CantidadesExistentes(SAPbouiCOM.Application sboapp, SAPbobsCOM.Company oCompany, SAPbouiCOM.Form oFormGL, ItemEvent pVal)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            try
            {
                SAPbouiCOM.Matrix oMatrixGL = (Matrix)oFormGL.Items.Item("mtxLO").Specific;

                #region Obtiene codigo y lote

                SAPbouiCOM.EditText otxtCA = (SAPbouiCOM.EditText)oFormGL.Items.Item("txtCA").Specific;
                SAPbouiCOM.EditText otxtWH = (SAPbouiCOM.EditText)oFormGL.Items.Item("txtWH").Specific;
                SAPbouiCOM.ComboBox cboDistNumber = (SAPbouiCOM.ComboBox)(oMatrixGL.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific);

                string stxtCA = null;
                string stxtWH = null;
                string scboDistNumber = null;
                string iQuantity = null;

                stxtCA = otxtCA.Value.ToString();
                stxtWH = otxtWH.Value.ToString();

                #endregion

                #region Valida si se ha selecciona el lote y consulta el lote

                if (cboDistNumber.Selected == null)
                {

                }
                else
                {
                    scboDistNumber = cboDistNumber.Selected.Value;

                    string sBachNumber = null;

                    SAPbobsCOM.Recordset oRBachNumber = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    sBachNumber = DllFunciones.GetStringXMLDocument(oCompany, "BOProduction", "Production", "GetQuantityBatchNumber");

                    sBachNumber = sBachNumber.Replace("%ItemCode%", stxtCA).Replace("%WhsCode%", stxtWH).Replace("%DistNumber%", scboDistNumber);

                    oRBachNumber.DoQuery(sBachNumber);

                    iQuantity = Convert.ToString(oRBachNumber.Fields.Item("Quantity").Value.ToString());

                    string sSepDecMachine = DllFunciones.GetSeparatorMachine();
                    string sSepDecSAP = DllFunciones.GetSeparatorSAP(oCompany);

                    if (sSepDecMachine == "," && sSepDecSAP == ".")
                    {
                        iQuantity = iQuantity.Replace(",", ".");
                    }
                    else if (sSepDecMachine == "." && sSepDecSAP == ".")
                    {
                        iQuantity = iQuantity.Replace(",", ".");
                    }

                    oFormGL.DataSources.UserDataSources.Item("DSCol0").Value = cboDistNumber.Selected.Value;
                    oFormGL.DataSources.UserDataSources.Item("DSCol1").ValueEx = iQuantity;

                    oMatrixGL.SetLineData(pVal.Row);

                    //SAPbouiCOM.Column oColumnCantidadDisponible = (SAPbouiCOM.Column)oMatrixGL.Columns.Item("Col_1");
                    //oColumnCantidadDisponible.ColumnSetting.SumType = BoColumnSumType.bst_Auto;

                    DllFunciones.liberarObjetos(oRBachNumber);

                }

                #endregion
            }
            catch (Exception)
            {

                throw;
            }


        }

        public Boolean Validate_WorkOrder(SAPbouiCOM.Application sboapp, SAPbobsCOM.Company oCompany, SAPbouiCOM.Form oFormNewWorkOrder)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            try
            {
                #region Variables y Objetos

                SAPbouiCOM.Matrix oMatrixNWO = (SAPbouiCOM.Matrix)oFormNewWorkOrder.Items.Item("mtxRP").Specific;

                string sCodArticulo = null;
                decimal iQuantity1 = 0;
                bool _BubbleEvent = false;

                #endregion

                #region Valida que no exista una linea duplicada o la cantidad esta en 0

                for (int i = 1; i <= oMatrixNWO.VisualRowCount; i++)
                {
                    sCodArticulo = ((SAPbouiCOM.EditText)(oMatrixNWO.Columns.Item("Col_1").Cells.Item(i).Specific)).Value;
                    iQuantity1 = Convert.ToDecimal(((SAPbouiCOM.EditText)(oMatrixNWO.Columns.Item("Col_3").Cells.Item(i).Specific)).Value);

                    if (iQuantity1 == 0)
                    {
                        DllFunciones.sendMessageBox(sboapp, "En el articulo " + sCodArticulo + " la cantidad esta en 0, por favor corrija para poder continuar.");

                        _BubbleEvent = false;
                        return _BubbleEvent;

                    }
                    else
                    {
                        #region Compara la Matriz buscando articulos duplicados

                        for (int j = i + 1; j <= oMatrixNWO.VisualRowCount; j++)
                        {
                            string sCodArticulo2 = null;

                            sCodArticulo2 = ((SAPbouiCOM.EditText)(oMatrixNWO.Columns.Item("Col_1").Cells.Item(j).Specific)).Value;

                            if (sCodArticulo == sCodArticulo2)
                            {
                                DllFunciones.sendMessageBox(sboapp, "El articulo " + sCodArticulo + " esta duplicado en las lineas de la orden de fabricacion, por favor corrija para poder continuar.");

                                _BubbleEvent = false;
                                return _BubbleEvent;
                            }
                        }

                        #endregion

                    }




                }

                int iContinuar = DllFunciones.sendMessageBoxY_N(sboapp, "Se creara la orden de produccion y sus semielaborados, ¿ Desea Continuar ?");

                if (iContinuar == 1)
                {
                    _BubbleEvent = true;
                    return _BubbleEvent;
                }
                else
                {
                    _BubbleEvent = false;
                    return _BubbleEvent;

                }

                #endregion

            }
            catch (Exception)
            {

                throw;
                return false;
            }

        }

        public Boolean Validate_BachtNumberTrasnfer(SAPbouiCOM.Application sboapp, SAPbobsCOM.Company oCompany, SAPbouiCOM.Form oFormNewWorkOrder, ItemEvent pVal)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            try
            {

                #region Variables y Objetos

                SAPbouiCOM.Matrix oMatrixLO = (SAPbouiCOM.Matrix)oFormNewWorkOrder.Items.Item("mtxLO").Specific;
                SAPbouiCOM.Matrix oMatrixLD = (SAPbouiCOM.Matrix)oFormNewWorkOrder.Items.Item("mtxLD").Specific;

                string iQuantity1 = null;
                decimal iSumQuantiy1 = 0;
                string iQuantity3 = null;
                decimal iSumQuantiy3 = 0;
                bool _BubbleEvent = false;
                string sDistNumber1;
                string sDistNumber3;

                CultureInfo CultureDecimal = new CultureInfo("en-US");

                string sSepDecMachine = DllFunciones.GetSeparatorMachine();
                string sSepDecSAP = DllFunciones.GetSeparatorSAP(oCompany);

                if (sSepDecMachine == "," && sSepDecSAP == ".")
                {
                    iSumQuantiy1 = 0;
                    iSumQuantiy3 = 0;
                }

                #endregion

                #region Valida que no exista una linea duplicada o la cantidad esta en 0 en la Matrix LO

                for (int i = 1; i <= oMatrixLO.VisualRowCount; i++)
                {

                    SAPbouiCOM.ComboBox cboDistNumber1 = (SAPbouiCOM.ComboBox)(oMatrixLO.Columns.Item("Col_0").Cells.Item(i).Specific);
                    sDistNumber1 = cboDistNumber1.Selected.Value;

                    iQuantity1 = ((SAPbouiCOM.EditText)(oMatrixLO.Columns.Item("Col_2").Cells.Item(i).Specific)).Value;

                    if (sSepDecMachine == "," && sSepDecSAP == ".")
                    {
                        iQuantity1 = iQuantity1.Replace(",", ".");
                    }

                    iSumQuantiy1 = Convert.ToDecimal(iQuantity1, CultureDecimal) + iSumQuantiy1;

                    if (iQuantity1 == "0" || iQuantity1 == "0.0" || iQuantity1 == "0.00" || iQuantity1 == "0.000" || iQuantity1 == "0.0000" || iQuantity1 == "0.00000" || iQuantity1 == "0.000000")
                    {
                        DllFunciones.sendMessageBox(sboapp, "En el Lote " + sDistNumber1 + " la cantidad esta en 0, por favor corrija para poder continuar.");

                        _BubbleEvent = false;
                        return _BubbleEvent;

                    }
                    else
                    {
                        #region Compara la Matriz buscando articulos duplicados

                        for (int j = i + 1; j <= oMatrixLO.VisualRowCount; j++)
                        {
                            SAPbouiCOM.ComboBox cboDistNumber2 = (SAPbouiCOM.ComboBox)(oMatrixLO.Columns.Item("Col_0").Cells.Item(j).Specific);
                            string sDistNumber2 = cboDistNumber2.Selected.Value;

                            sDistNumber2 = cboDistNumber2.Selected.Value;

                            if (sDistNumber1 == sDistNumber2)
                            {
                                DllFunciones.sendMessageBox(sboapp, "El Lote " + sDistNumber1 + " esta duplicado en las lineas de los lotes a retirar, por favor corrija para poder continuar.");

                                _BubbleEvent = false;
                                return _BubbleEvent;
                            }
                        }

                        #endregion
                    }
                }

                #endregion

                #region Valida que no exista una linea duplicada o la cantidad esta en 0 en la Matrix LD

                for (int i = 1; i <= oMatrixLD.VisualRowCount; i++)
                {
                    sDistNumber3 = ((SAPbouiCOM.EditText)(oMatrixLD.Columns.Item("Col_0").Cells.Item(i).Specific)).Value;

                    if (string.IsNullOrEmpty(sDistNumber3))
                    {
                        DllFunciones.sendMessageBox(sboapp, "No se ha ingresado ningun lote a recibir, por favor corrija para continuar");
                        _BubbleEvent = false;
                        return _BubbleEvent;
                    }
                    else
                    {
                        iQuantity3 = ((SAPbouiCOM.EditText)(oMatrixLD.Columns.Item("Col_2").Cells.Item(i).Specific)).Value;

                        if (sSepDecMachine == "," && sSepDecSAP == ".")
                        {
                            iQuantity3 = iQuantity3.Replace(",", ".");
                        }

                        iSumQuantiy3 = Convert.ToDecimal(iQuantity3, CultureDecimal) + iSumQuantiy3;

                        if (iQuantity3 == "0" || iQuantity3 == "0.0" || iQuantity3 == "0.00" || iQuantity3 == "0.000" || iQuantity3 == "0.0000" || iQuantity3 == "0.00000" || iQuantity3 == "0.000000")
                        {
                            DllFunciones.sendMessageBox(sboapp, "En el Lote " + sDistNumber3 + " la cantidad esta en 0, por favor corrija para poder continuar.");

                            _BubbleEvent = false;
                            return _BubbleEvent;

                        }
                        else
                        {
                            #region Compara la Matriz buscando articulos duplicados

                            for (int j = i + 1; j <= oMatrixLD.VisualRowCount; j++)
                            {

                                string sDistNumber4 = ((SAPbouiCOM.EditText)(oMatrixLD.Columns.Item("Col_0").Cells.Item(j).Specific)).Value;

                                if (sDistNumber3 == sDistNumber4)
                                {
                                    DllFunciones.sendMessageBox(sboapp, "El Lote " + sDistNumber3 + " esta duplicado en las lineas de los lotes a retirar, por favor corrija para poder continuar.");

                                    _BubbleEvent = false;
                                    return _BubbleEvent;
                                }
                            }

                            #endregion
                        }
                    }
                }
                #endregion

                #region Validacion cantidades a enviar y reibir

                if (iSumQuantiy1 != iSumQuantiy3)
                {
                    DllFunciones.sendMessageBox(sboapp, "Las cantidades a retirar son diferentes a las cantidad a ingresar, por favor igualar las cantidades en los lotes, para poder continuar");

                    _BubbleEvent = false;
                    return _BubbleEvent;

                }

                #endregion

                int iContinuar = DllFunciones.sendMessageBoxY_N(sboapp, "Se realizara la transferencia de lotes, ¿ Desea Continuar ?");

                if (iContinuar == 1)
                {
                    _BubbleEvent = true;
                    return _BubbleEvent;
                }
                else
                {
                    _BubbleEvent = false;
                    return _BubbleEvent;

                }
            }
            catch (Exception)
            {

                throw;
                return false;
            }

        }

        public void CreateGoodIssueandGoodReceipt(SAPbouiCOM.Application sboapp, SAPbobsCOM.Company oCompany, SAPbouiCOM.Form oFormGL)
        {
            Funciones.Comunes DllFuciones = new Funciones.Comunes();
            try
            {
                DllFuciones.sendStatusBarMsg(sboapp, "Transfiriendo lote, por espere...", BoMessageTime.bmt_Medium, false);

                #region Variables y objetos

                SAPbouiCOM.Matrix oMatrixLO = (SAPbouiCOM.Matrix)oFormGL.Items.Item("mtxLO").Specific;
                SAPbouiCOM.Matrix oMatrixLD = (SAPbouiCOM.Matrix)oFormGL.Items.Item("mtxLD").Specific;

                SAPbouiCOM.EditText otxtCA = (SAPbouiCOM.EditText)oFormGL.Items.Item("txtCA").Specific;
                SAPbouiCOM.EditText otxtWH = (SAPbouiCOM.EditText)oFormGL.Items.Item("txtWH").Specific;

                SAPbobsCOM.Recordset oRNumberSeriesActive = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset oRConsecutivoTL = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string sNumberSeriesActive;
                string sConsecutivoTL;
                decimal iSumQuantityIssue = 0;
                int Rsd = 0;

                string sComments = "Documento realizado automaticamente por gestor de lotes";

                CultureInfo CultureDecimal = new CultureInfo("en-US");

                #endregion

                #region Busqueda de series de numeracion asignada y cuenta contable compensacion

                sNumberSeriesActive = DllFuciones.GetStringXMLDocument(oCompany, "BOProduction", "Production", "GetNumberSerieActive");

                oRNumberSeriesActive.DoQuery(sNumberSeriesActive);

                #endregion

                #region Consulta consecutivo transferencia de lote

                sConsecutivoTL = DllFuciones.GetStringXMLDocument(oCompany, "BOProduction", "Production", "GetNextDocEntryTL");

                oRConsecutivoTL.DoQuery(sConsecutivoTL);

                #endregion

                if (string.IsNullOrEmpty(oRNumberSeriesActive.Fields.Item("Code_SNEM").Value.ToString()) || string.IsNullOrEmpty(oRNumberSeriesActive.Fields.Item("Code_SNSM").Value.ToString()) || string.IsNullOrEmpty(oRNumberSeriesActive.Fields.Item("Cuenta_Compensacion").Value.ToString()))
                {
                    DllFuciones.sendMessageBox(sboapp, "No se puede realizar la transferencia ya que no se han configurado los parametros iniciales, por favor validar con el adminsitrador del sistema");
                }
                else
                {
                    #region Creacion de la salida de mercancia

                    #region Encabezado

                    SAPbobsCOM.Documents DocumentGoodIssue = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);

                    DocumentGoodIssue.Series = Convert.ToInt32(oRNumberSeriesActive.Fields.Item("Code_SNSM").Value.ToString());
                    DocumentGoodIssue.DocDate = DateTime.Now;
                    DocumentGoodIssue.TaxDate = DateTime.Now;
                    DocumentGoodIssue.Comments = sComments;
                    DocumentGoodIssue.Reference2 = Convert.ToString(oRConsecutivoTL.Fields.Item("Consecutivo").Value.ToString());

                    #endregion

                    #region Lineas

                    DocumentGoodIssue.Lines.ItemCode = otxtCA.Value.ToString();

                    for (int i = 1; i <= oMatrixLO.VisualRowCount; i++)
                    {
                        string iCantidadaSumar = ((SAPbouiCOM.EditText)(oMatrixLO.Columns.Item("Col_2").Cells.Item(i).Specific)).Value;
                        iCantidadaSumar = iCantidadaSumar.Replace(".", ",");
                        iSumQuantityIssue = iSumQuantityIssue + Convert.ToDecimal(iCantidadaSumar, CultureDecimal);
                    }

                    DocumentGoodIssue.Lines.Quantity = Convert.ToDouble(iSumQuantityIssue);
                    DocumentGoodIssue.Lines.WarehouseCode = otxtWH.Value.ToString();
                    DocumentGoodIssue.Lines.AccountCode = Convert.ToString(oRNumberSeriesActive.Fields.Item("Cuenta_Compensacion").Value.ToString());

                    for (int i = 1; i <= oMatrixLO.VisualRowCount; i++)
                    {
                        #region Asignacion de lotes

                        DocumentGoodIssue.Lines.BatchNumbers.BaseLineNumber = i - 1;

                        SAPbouiCOM.ComboBox cboBatchNumber = (SAPbouiCOM.ComboBox)(oMatrixLO.Columns.Item("Col_0").Cells.Item(i).Specific);
                        string sBatchNumber = cboBatchNumber.Selected.Value;

                        DocumentGoodIssue.Lines.BatchNumbers.BatchNumber = sBatchNumber;

                        string iCantidadLote = ((SAPbouiCOM.EditText)(oMatrixLO.Columns.Item("Col_2").Cells.Item(i).Specific)).Value;
                        iCantidadLote = iCantidadLote.Replace(".", ",");

                        DocumentGoodIssue.Lines.BatchNumbers.Quantity = Convert.ToDouble(iCantidadLote);
                        DocumentGoodIssue.Lines.BatchNumbers.Add();

                        #endregion
                    }

                    DocumentGoodIssue.Lines.Add();

                    #endregion

                    Rsd = DocumentGoodIssue.Add();

                    if (Rsd != 0)
                    {
                        DllFuciones.sendMessageBox(sboapp, "Error: " + Convert.ToString(oCompany.GetLastErrorCode()) + " : " + oCompany.GetLastErrorDescription());
                        DllFuciones.liberarObjetos(DocumentGoodIssue);
                        DllFuciones.liberarObjetos(oRNumberSeriesActive);
                        DllFuciones.liberarObjetos(oRConsecutivoTL);
                    }
                    else
                    {
                        int sDocEntryGoodIssue = Convert.ToInt32(oCompany.GetNewObjectKey());

                        #region Variables y objetos

                        SAPbobsCOM.Recordset oRInfoGoodIssue = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                        string sSearchGoodIssue = null;
                        decimal iSumQuantityGoodReceipt = 0;

                        Rsd = 0;

                        #endregion

                        #region Consulta el costo de la salida

                        sSearchGoodIssue = DllFuciones.GetStringXMLDocument(oCompany, "BOProduction", "Production", "GetInfoGoodIssue");

                        sSearchGoodIssue = sSearchGoodIssue.Replace("%DocEntry%", Convert.ToString(sDocEntryGoodIssue));

                        oRInfoGoodIssue.DoQuery(sSearchGoodIssue);

                        #endregion

                        #region Creacion de la entrada de Mercancia

                        #region Encabezado

                        SAPbobsCOM.Documents GoodReceipt = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);

                        GoodReceipt.Series = Convert.ToInt32(oRNumberSeriesActive.Fields.Item("Code_SNEM").Value.ToString());
                        GoodReceipt.DocDate = DateTime.Now;
                        GoodReceipt.TaxDate = DateTime.Now;
                        GoodReceipt.Reference2 = Convert.ToString(oRConsecutivoTL.Fields.Item("Consecutivo").Value.ToString());
                        GoodReceipt.Comments = sComments;

                        #endregion

                        #region Lineas

                        GoodReceipt.Lines.ItemCode = otxtCA.Value.ToString();

                        for (int i = 1; i <= oMatrixLD.VisualRowCount; i++)
                        {
                            string xCantidadaSumar = ((SAPbouiCOM.EditText)(oMatrixLD.Columns.Item("Col_2").Cells.Item(i).Specific)).Value;
                            xCantidadaSumar = xCantidadaSumar.Replace(".", ",");
                            iSumQuantityGoodReceipt = iSumQuantityGoodReceipt + Convert.ToDecimal(xCantidadaSumar, CultureDecimal);

                        }

                        GoodReceipt.Lines.Quantity = Convert.ToDouble(iSumQuantityGoodReceipt);
                        GoodReceipt.Lines.UnitPrice = Convert.ToDouble(oRInfoGoodIssue.Fields.Item("StockPrice").Value.ToString());
                        GoodReceipt.Lines.WarehouseCode = otxtWH.Value.ToString();
                        GoodReceipt.Lines.AccountCode = Convert.ToString(oRNumberSeriesActive.Fields.Item("Cuenta_Compensacion").Value.ToString());

                        for (int i = 1; i <= oMatrixLD.VisualRowCount; i++)
                        {
                            #region Asignacion de lotes

                            GoodReceipt.Lines.BatchNumbers.BaseLineNumber = i - 1;
                            GoodReceipt.Lines.BatchNumbers.BatchNumber = Convert.ToString(((SAPbouiCOM.EditText)(oMatrixLD.Columns.Item("Col_0").Cells.Item(i).Specific)).Value);

                            string xCantidadLote = ((SAPbouiCOM.EditText)(oMatrixLO.Columns.Item("Col_2").Cells.Item(i).Specific)).Value;
                            xCantidadLote = xCantidadLote.Replace(".", ",");

                            GoodReceipt.Lines.BatchNumbers.Quantity = Convert.ToDouble(xCantidadLote, CultureDecimal);
                            GoodReceipt.Lines.BatchNumbers.Add();

                            #endregion
                        }

                        GoodReceipt.Lines.Add();

                        #endregion

                        Rsd = GoodReceipt.Add();

                        if (Rsd != 0)
                        {
                            DllFuciones.sendMessageBox(sboapp, "Error: " + Convert.ToString(oCompany.GetLastErrorCode()) + " : " + oCompany.GetLastErrorDescription());

                            #region Liberar objetos

                            DllFuciones.liberarObjetos(DocumentGoodIssue);
                            DllFuciones.liberarObjetos(GoodReceipt);
                            DllFuciones.liberarObjetos(oRNumberSeriesActive);
                            DllFuciones.liberarObjetos(oRConsecutivoTL);
                            DllFuciones.liberarObjetos(oRInfoGoodIssue);

                            #endregion

                        }
                        else
                        {
                            int sDocEntryGoodRecepit = Convert.ToInt32(oCompany.GetNewObjectKey());

                            #region Actualiza tabla gestion de lotes

                            SAPbobsCOM.UserTable oUDT_TL;

                            oUDT_TL = (SAPbobsCOM.UserTable)(oCompany.UserTables.Item("BOTL"));

                            oUDT_TL.Code = Convert.ToString(oRConsecutivoTL.Fields.Item("Consecutivo").Value.ToString());
                            oUDT_TL.Name = Convert.ToString(oRConsecutivoTL.Fields.Item("Consecutivo").Value.ToString());

                            oUDT_TL.UserFields.Fields.Item("U_BO_ItemCode").Value = otxtCA.Value.ToString();
                            oUDT_TL.UserFields.Fields.Item("U_BO_WhsCode").Value = otxtWH.Value.ToString();
                            oUDT_TL.UserFields.Fields.Item("U_BO_DocEntryOIGE").Value = Convert.ToString(sDocEntryGoodIssue);
                            oUDT_TL.UserFields.Fields.Item("U_BO_DocEntryOIGN").Value = Convert.ToString(sDocEntryGoodRecepit);

                            Rsd = oUDT_TL.Add();

                            #endregion

                            #region Liberar Objetos

                            DllFuciones.liberarObjetos(DocumentGoodIssue);
                            DllFuciones.liberarObjetos(GoodReceipt);
                            DllFuciones.liberarObjetos(oRNumberSeriesActive);
                            DllFuciones.liberarObjetos(oRConsecutivoTL);
                            DllFuciones.liberarObjetos(oRInfoGoodIssue);


                            #endregion

                            DllFuciones.StatusBar(sboapp, BoStatusBarMessageType.smt_Success, "Operacion finalizada con exito");

                            DllFuciones.sendMessageBox(sboapp, "Se transfirio el lote correctamente");

                            DllFuciones.CloseFormXML(sboapp, "BO_GL");


                        }

                        #endregion


                    }

                    #endregion
                }
            }
            catch (Exception)
            {

                throw;
            }

        }

        public void Right_Click(ref SAPbouiCOM.ContextMenuInfo _eventInfo, SAPbouiCOM.Application _sboapp)
        {
            SAPbouiCOM.MenuItem oMenuItem = null;
            SAPbouiCOM.Menus oMenus = null;

            try
            {
                #region Click derecho para adicionar Servicio de teñido

                if (_sboapp.Menus.Exists("AddOSM"))
                {
                   // _sboapp.Menus.RemoveEx("AddOSM");
                }
                else
                {
                    SAPbouiCOM.MenuCreationParams oCreationPackage = null;
                    oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(_sboapp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));

                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    oCreationPackage.UniqueID = "AddOSM";
                    oCreationPackage.String = "Crear OS Maquila";
                    oCreationPackage.Enabled = true;
                    oCreationPackage.Image = (Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\BOProduction\\Images\\UpArrow.bmp");

                    oMenuItem = _sboapp.Menus.Item("1280"); // Data'
                    oMenus = oMenuItem.SubMenus;
                    oMenus.AddEx(oCreationPackage);

                    #endregion

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
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
