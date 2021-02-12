using SAPbouiCOM;
using SAPbobsCOM;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml;
using System.Windows.Forms;


namespace Funciones
{
    public class Comunes
    {

        private int lRetCode;
        SAPbouiCOM.ProgressBar oProgressBar = null;
        int Rsd = 0;
        int Contador = 0;
        string sPath;
        string sMotor;
        string sNameDB;
        string sSearchCategory;

        public void AddFormatedSearch(SAPbobsCOM.Company _oCompany, SAPbouiCOM.Application _sboapp, string _CategoryName, string _NameSearchFormatted, string _DllName, string _IDNodo1, string _IDNodo2)
        {
            //SAPbobsCOM.FormattedSearches oFormatted = null;
            string CategoryID = null;

            CategoryID = SearchCatetoryID(_oCompany, _CategoryName, _DllName);

            if (CategoryID == "0")
            {
                #region Crea Categoria si no existe y retorna el ID

                CategoryID = AddCategory(_oCompany, _sboapp, _CategoryName);

                #endregion

                #region Adiciona la consulta al Query Manager

                AddQueryManager(_oCompany, _sboapp, CategoryID, _NameSearchFormatted, _DllName, _IDNodo1, _IDNodo2);

                #endregion

            }
            else
            {
                #region Adiciona la consulta al Query Manager

                AddQueryManager(_oCompany, _sboapp, CategoryID, _NameSearchFormatted, _DllName, _IDNodo1, _IDNodo2);

                #endregion

            }

        }

        private string AddCategory(SAPbobsCOM.Company _oCompany, SAPbouiCOM.Application _sboapp, string NameCategory)
        {
            string NewCategory = null;
            int Rsd;

            SAPbobsCOM.QueryCategories oCategory;
            oCategory = (SAPbobsCOM.QueryCategories)_oCompany.GetBusinessObject(BoObjectTypes.oQueryCategories);
            oCategory.Name = NameCategory;

            Rsd = oCategory.Add();

            if (Rsd == 0)
            {
                NewCategory = _oCompany.GetNewObjectKey();
                return NewCategory;
                liberarObjetos(oCategory);
            }
            else
            {
                sendMessageBox(_sboapp, "No se pudo crear la categoria eBilling");
                return NewCategory;
            }
            liberarObjetos(oCategory);

        }

        public void AddQueryManager(SAPbobsCOM.Company _oCompany, SAPbouiCOM.Application _sboapp, string _CategoryID, string _NameSearchFormatted, string _DllName, string _IDNodo1, string _IDNodo2)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            string sQueryActividadEconomica = null;
            string sCategoryInternalKey = null;
            Rsd = 0;


            sPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            sMotor = Convert.ToString(_oCompany.DbServerType);

            #region Busca la busqueda formateada en el Query Manager

            sSearchCategory = GetStringXMLDocument(_oCompany, _DllName, _DllName, _IDNodo1);
            sSearchCategory = sSearchCategory.Replace("%NameSearchFormatted%", _NameSearchFormatted).Replace("%CategoryID%", _CategoryID);

            SAPbobsCOM.UserQueries oCreateQueryManager = (SAPbobsCOM.UserQueries)(_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserQueries));
            SAPbobsCOM.Recordset oSearchFormatted = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            oSearchFormatted.DoQuery(sSearchCategory);

            #endregion

            sCategoryInternalKey = oSearchFormatted.Fields.Item(0).Value.ToString();

            sQueryActividadEconomica = GetStringXMLDocument(_oCompany, _DllName, _DllName, _IDNodo2);

            if (sCategoryInternalKey == "0")
            {
                #region Crea la busqueda formateada en el Query Manager

                oCreateQueryManager.Query = sQueryActividadEconomica;
                oCreateQueryManager.QueryCategory = Convert.ToInt32(_CategoryID);
                oCreateQueryManager.QueryDescription = _NameSearchFormatted;

                Rsd = oCreateQueryManager.Add();

                if (Rsd == 0)
                {

                    liberarObjetos(oCreateQueryManager);
                    liberarObjetos(oSearchFormatted);
                }
                else
                {
                    DllFunciones.sendMessageBox(_sboapp, "No se puedo crear la busqueda formateada");
                }
                #endregion
            }
            else
            {
                #region Actualiza la busqueda formateada 

                oCreateQueryManager.GetByKey(Convert.ToInt32(sCategoryInternalKey), Convert.ToInt32(_CategoryID));
                oCreateQueryManager.Query = sQueryActividadEconomica;
                oCreateQueryManager.QueryCategory = Convert.ToInt32(_CategoryID);
                oCreateQueryManager.QueryDescription = _NameSearchFormatted;

                Rsd = oCreateQueryManager.Update();

                if (Rsd == 0)
                {
                    liberarObjetos(oCreateQueryManager);
                    liberarObjetos(oSearchFormatted);
                }
                else
                {
                    DllFunciones.sendMessageBox(_sboapp, _oCompany.GetLastErrorDescription());
                }

                #endregion
            }
        }

        /// <summary>
        /// Método para cargar formulario por XML
        /// </summary>
        /// <param name="sboapp">Objeto oAplication</param>
        /// <param name="stream">objeto stream para cargue de formulario XML</param>
        /// <returns></returns>    
        public Boolean crearFormPorXML(SAPbouiCOM.Application sboapp, System.IO.Stream stream)
        {
            try
            {
                //String frmResource = $"{Assembly.GetExecutingAssembly().GetName().Name}.Formularios.ConfPresup.srf";
                //System.IO.Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(frmResource);
                StreamReader sr = new StreamReader(stream);
                String formXml = sr.ReadToEnd();
                sboapp.LoadBatchActions(formXml);
                return true;
            }
            catch (Exception e)
            {
                sendErrorMessage(sboapp, e);
                return false;
            }

        }

        /// <summary>
        /// Método para crear tablas de usuario
        /// </summary>
        /// <param name="oCompany">objeto Company del DI-API SAPBobscom seteado en _company</param>
        /// <param name="Name">Texto con el nombre de la tabla</param>
        /// <param name="Description">Texto con la descripcion de la tabla</param>
        /// <param name="Type">Tipo de tabla BoUTBTableType.bott_Document  Documentos, Lineas, Datos maestros o ninguno</param>
        /// <returns></returns>
        public Boolean crearTabla(SAPbobsCOM.Company oCompany, SAPbouiCOM.Application sboapp, string Name, string Description, SAPbobsCOM.BoUTBTableType Type)
        {
            SAPbobsCOM.UserTablesMD oUserTablesMD = null;
            try
            {

                oUserTablesMD = ((SAPbobsCOM.UserTablesMD)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)));
                if (false.Equals(oUserTablesMD.GetByKey(Name)))
                {
                    oUserTablesMD.TableName = Name;
                    oUserTablesMD.TableDescription = Description;
                    oUserTablesMD.TableType = Type;
                    lRetCode = oUserTablesMD.Add();
                    if (lRetCode != 0)
                    {
                        if (lRetCode == -1)
                        {
                            return true;
                        }
                        else
                        {
                            sendMessageBox(sboapp, $"Incidente: {lRetCode}");
                            return false;
                        }
                    }
                    else
                    {
                        return true;
                    }
                }
                else
                {
                    //Ya existe la tabla
                    return true;
                }

            }
            catch (Exception e)
            {
                sendErrorMessage(sboapp, e);
                return false;
            }
            finally
            {
                liberarObjetos(oUserTablesMD);
            }
        }

        /// <summary>
        /// Método que cierra el formulario
        /// </summary>
        public void CloseFormXML(SAPbouiCOM.Application sboapp, string UIDForm)
        {
            SAPbouiCOM.Form oForm;
            oForm = sboapp.Forms.Item(UIDForm);
            oForm.Close();
            liberarObjetos(oForm);
        }

        /// <summary>
        /// Método para crear un UDO con tablas de detalle
        /// </summary>
        /// <param name="oCompany"></param>
        /// <param name="oAplication"></param>
        /// <param name="Code">Código interno del UDO</param>
        /// <param name="ObjName">Descripción del UDO</param>
        /// <param name="ObjectType">Código para objtype</param>
        /// <param name="Tables">arreglo que contiene todas las tablas inicia con la de encabezado</param>
        /// <param name="CanDelete">se permite borrar</param>
        /// <param name="CanFind">se permite buscar</param>
        /// <param name="FindColumns">columnas por las cuales se permite la búsqueda</param>
        /// <param name="CanCancel">se permite cancelar</param>
        /// <param name="CanClose">se permite cerrar</param>
        /// <param name="ManageSeries">maneja series de SAP</param>
        /// <param name="CanYearTransfer">se permite traslado de año</param>
        /// <param name="CanCreateDefaultForm"></param>
        /// <param name="EnableEnhancedForm"></param>
        /// <param name="UseUniqueFormType"></param>
        /// <param name="MenuItem">menu en el que se desea ubicar</param>
        /// <param name="Position">posicion dentro del menu</param>
        /// <param name="FatherMenuID"></param>
        /// <param name="CanLog">permite registrar en log</param>
        /// <param name="LogTableName">Nombre de la tabla para registrar el log</param>
        /// <returns></returns>
        public Boolean CrearUDO(SAPbobsCOM.Company oCompany, SAPbouiCOM.Application oAplication, String Code, String ObjName, SAPbobsCOM.BoUDOObjType ObjectType, String[] Tables, SAPbobsCOM.BoYesNoEnum CanDelete, BoYesNoEnum CanFind, String[] FindColumns, BoYesNoEnum CanCancel, BoYesNoEnum CanClose, BoYesNoEnum ManageSeries, BoYesNoEnum CanYearTransfer, BoYesNoEnum CanCreateDefaultForm, BoYesNoEnum EnableEnhancedForm, BoYesNoEnum UseUniqueFormType, BoYesNoEnum MenuItem, int Position, int FatherMenuID, BoYesNoEnum CanLog, String LogTableName)
        {
            SAPbobsCOM.UserObjectsMD oUdtMD;

            try
            {
                oUdtMD = (SAPbobsCOM.UserObjectsMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);

                if (oUdtMD.GetByKey(Code) == false)
                {
                    oUdtMD.Code = Code;
                    oUdtMD.Name = ObjName;
                    oUdtMD.TableName = (string)Tables.GetValue(0);

                    if (Tables.Length > 1)
                    {
                        for (int i = 1; i < Tables.Length; i++)
                        {
                            oUdtMD.ChildTables.TableName = (string)Tables.GetValue(i);
                            oUdtMD.ChildTables.Add();
                        }
                    }
                    oUdtMD.ObjectType = ObjectType;
                    oUdtMD.CanDelete = CanDelete;
                    oUdtMD.CanFind = CanFind;
                    oUdtMD.CanCancel = CanCancel;
                    oUdtMD.CanClose = CanClose;
                    oUdtMD.ManageSeries = ManageSeries;
                    oUdtMD.CanYearTransfer = CanYearTransfer;
                    oUdtMD.CanCreateDefaultForm = CanCreateDefaultForm;
                    oUdtMD.EnableEnhancedForm = EnableEnhancedForm;

                    if (oUdtMD.CanCreateDefaultForm == BoYesNoEnum.tYES || oUdtMD.EnableEnhancedForm == SAPbobsCOM.BoYesNoEnum.tYES)
                    {
                        oUdtMD.UseUniqueFormType = UseUniqueFormType;

                        if (ObjectType == SAPbobsCOM.BoUDOObjType.boud_MasterData)
                        {
                            oUdtMD.FormColumns.FormColumnAlias = "Code";
                            oUdtMD.FormColumns.Add();
                        }
                        else
                        {
                            oUdtMD.FormColumns.FormColumnAlias = "DocEntry";
                            oUdtMD.FormColumns.Add();
                        }
                    }
                    oUdtMD.MenuItem = MenuItem;

                    if (oUdtMD.MenuItem == BoYesNoEnum.tYES)
                    {
                        oUdtMD.Position = Position;
                        oUdtMD.FatherMenuID = FatherMenuID;
                    }
                    oUdtMD.CanLog = CanLog;

                    if (oUdtMD.CanLog == BoYesNoEnum.tYES)
                    {
                        oUdtMD.LogTableName = LogTableName;
                    }

                    if (CanFind == BoYesNoEnum.tYES)
                    {
                        int Fields = 0;
                        try
                        {
                            Fields = FindColumns.Length;
                        }
                        catch (Exception ex)
                        {

                            Fields = 0;

                        }
                        if (Fields > 0)
                        {
                            for (int i = 0; i < FindColumns.Length - 1; i++)
                            {
                                oUdtMD.FindColumns.ColumnAlias = (string)FindColumns.GetValue(i);
                                oUdtMD.FindColumns.Add();
                            }
                        }
                        else
                        {
                            if (ObjectType == BoUDOObjType.boud_MasterData)
                            {
                                oUdtMD.FindColumns.ColumnAlias = "Code";
                                oUdtMD.FindColumns.Add();
                            }
                            else
                            {
                                oUdtMD.FindColumns.ColumnAlias = "DocEntry";
                                oUdtMD.FindColumns.Add();
                            }
                        }
                    }

                    if (oUdtMD.Add() != 0)
                    {
                        sendMessageBox(oAplication, $"Incidente: {oCompany.GetLastErrorDescription()}", 1);
                        return false;
                    }

                }
                return true;
            }
            catch (Exception ex)
            {
                sendErrorMessage(oAplication, ex);
                return false;
            }
            finally
            {

            }
            liberarObjetos(oUdtMD);
        }

        /// <summary>
        /// Método para crear campos de usuario
        /// </summary>
        /// <param name="oCompany">Parámetro de Compañía Company</param>
        /// <param name="oAplication">Parámetro de Aplicación Application</param>
        /// <param name="FieldType">Tipo de campo BoFieldTypes</param>
        /// <param name="FieldSubType">Subtipo de campo BoFldSubTypes</param>
        /// <param name="EditSize">Tamaño del campo</param>
        /// <param name="DfltValue">Valor por defecto </param>
        /// <param name="Mandatory">Campo obligatorio BoYesNoEnum True o false</param>
        /// <param name="ValidValues">Arreglo con los valores válidos</param>
        /// <param name="TableName">Nombre de la tabla</param>
        /// <param name="FieldName">Nombre del campo</param>
        /// <param name="FieldDescription">Descripción del campo</param>
        /// <returns></returns>
        public Boolean CreaCamposUsr(SAPbobsCOM.Company oCompany, SAPbouiCOM.Application oAplication, SAPbobsCOM.BoFieldTypes FieldType, SAPbobsCOM.BoFldSubTypes FieldSubType, int EditSize, string DfltValue, SAPbobsCOM.BoYesNoEnum Mandatory, string[] ValidValues, string TableName, string FieldName, string FieldDescription)
        {

            int Nro = 0;
            SAPbobsCOM.UserFieldsMD FieldMD;
            string Message = "";
            Boolean AddCode = true;
            Boolean AddWVV = true;
            int Add = 0;
            Boolean NeedToUpDate = false;

            try
            {
                FieldMD = (SAPbobsCOM.UserFieldsMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                try
                {
                    if (ValidValues[0] != "")
                    {
                        int c = ValidValues.Count();
                        if (c > 0)
                        {
                            Nro = ValidValues.Count();
                        }
                    }

                    else
                    {
                        Nro = 0;
                    }
                }
                catch (Exception)
                {
                    Nro = 0;
                }

                if (TableName != "")
                {
                    if (FieldName != "")
                    {
                        if (FieldDescription != "")
                        {
                            bool bUpdate;
                            // 
                            //bUpdate = FieldMD.GetByKey(TableName, Convert.ToInt32(FieldName));
                            bUpdate = FieldMD.GetByKey(TableName, FieldSecuence(oCompany, oAplication, TableName, FieldName));

                            if (bUpdate == false)
                            {
                                FieldMD.TableName = TableName;
                                FieldMD.Name = FieldName;
                                FieldMD.Description = FieldDescription;
                                FieldMD.Type = FieldType;
                                if (FieldSubType != SAPbobsCOM.BoFldSubTypes.st_None)
                                {
                                    FieldMD.SubType = FieldSubType;
                                }
                                if (EditSize != 0)
                                {
                                    FieldMD.EditSize = EditSize;
                                }
                                if (DfltValue != "")
                                {
                                    FieldMD.DefaultValue = DfltValue;
                                }
                                if (Mandatory != SAPbobsCOM.BoYesNoEnum.tNO)
                                {
                                    FieldMD.Mandatory = Mandatory;
                                }
                                if (Nro != 0)
                                {
                                    if (Nro % 2 != 0)
                                    {
                                        Message = "El arreglo de valores valido debe tener un número par de datos";
                                    }
                                    foreach (var Val in ValidValues)
                                    {
                                        if (Val != "")
                                        {
                                            if (AddCode)
                                            {
                                                FieldMD.ValidValues.Value = Val;
                                                AddCode = false;
                                            }
                                            else
                                            {
                                                FieldMD.ValidValues.Description = Val;
                                                AddCode = true;
                                            }
                                            if (AddCode)
                                            {
                                                FieldMD.ValidValues.Add();
                                            }
                                        }
                                        else
                                        {
                                            if (AddCode)
                                            {
                                                Message = "El código no puede estar vacío";
                                                AddWVV = false;
                                                break;
                                            }
                                            else
                                            {
                                                Message = "La descripción no puede estar vacía";
                                                AddWVV = false;
                                                break;
                                            }
                                        }

                                    }
                                }
                                if (AddWVV)
                                {
                                    Add = FieldMD.Add();
                                    if (Add != 0)
                                    {
                                        oCompany.GetLastError(out Add, out Message);
                                        if (Message != "Ref count for this object is higher then 0")
                                        {
                                            if (Message != "Esta entrada ya existe en las tablas siguientes (ODBC -2035)")
                                            {
                                                if (Message != "This entry already exists in the following tables (ODBC -2035)")
                                                {
                                                    if (Message.Trim() != "No records")
                                                    {
                                                        if (Message.Trim() != "Sin registros")
                                                        {
                                                            Message = "Error al crear el campo: " + FieldName + " en la tabla: " + TableName + ", " + Message;

                                                        }
                                                        else
                                                        {
                                                            Message = "";
                                                        }
                                                    }
                                                    else
                                                    {
                                                        Message = "";
                                                    }
                                                }
                                                else
                                                {
                                                    Message = "";
                                                }
                                            }
                                            else
                                            {
                                                Message = "";
                                            }
                                        }
                                    }
                                }

                            }
                            else
                            {
                                bool bUpdate2;
                                // 
                                //bUpdate = FieldMD.GetByKey(TableName, Convert.ToInt32(FieldName));
                                bUpdate2 = FieldMD.GetByKey(TableName, FieldSecuence(oCompany, oAplication, TableName, FieldName));
                                if (bUpdate2)
                                {
                                    if (FieldMD.Description != FieldDescription)
                                    {
                                        FieldMD.Description = FieldDescription;
                                        NeedToUpDate = true;
                                    }
                                    if (FieldMD.EditSize != EditSize)
                                    {
                                        FieldMD.EditSize = EditSize;
                                        NeedToUpDate = true;
                                    }
                                    if (FieldMD.DefaultValue != DfltValue)
                                    {
                                        FieldMD.DefaultValue = DfltValue;
                                        NeedToUpDate = true;
                                    }
                                    if (NeedToUpDate)
                                    {
                                        Add = FieldMD.Update();
                                    }
                                }
                            }

                        }
                    }

                }
                liberarObjetos(FieldMD);
                if (Message != "")
                {
                    sendMessageBox(oAplication, Message, 1);
                    return false;
                }
                else
                {
                    return true;
                }
            }

            catch (Exception ex)
            {
                sendErrorMessage(oAplication, ex);
                return false;
            }
        }

        /// <summary>
        /// Método para importar los datos de un archivo archivos CSV
        /// </summary>
        public void ImportCSV(SAPbouiCOM.Application _sboapp, SAPbobsCOM.Company _oCompany, string NombreArchivoCSV, string DllName, string _IDNodoTabla, string IDNodoInsert)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();
            try
            {
                #region Variables y objetos

                string sLinea = null;
                string[] sArreglo;
                string sCode = null;
                string sName = null;
                string sDescripcion = null;
                string sInsertarOriginal = null;
                string sInsertarProcesado = null;
                string sVAlidacion = null;
                int sCantidadColumnas = 0;
                int sSalir = 0;

                SAPbobsCOM.Recordset oInsertar = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset oValidacion = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                #endregion

                #region Validacion si existen datos en la tabla

                sVAlidacion = DllFunciones.GetStringXMLDocument(_oCompany, "eBilling", "eBilling", _IDNodoTabla);

                oValidacion.DoQuery(sVAlidacion);

                #endregion

                #region Asignacion cantidad de Columnas del archivo

                if (NombreArchivoCSV == "Tiposresponsabilidades")
                {
                    sCantidadColumnas = 3;
                }
                else if (NombreArchivoCSV == "UnidadesdeMedidaDIAN")
                {
                    sCantidadColumnas = 2;
                }
                else
                {
                    sCantidadColumnas = 2;
                }

                #endregion

                if (oValidacion.RecordCount == 0)
                {
                    #region Importa los datos del archivos CSV

                    sPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

                    sPath = sPath + "\\" + DllName + "\\CSV\\" + NombreArchivoCSV + ".csv";

                    var ArchivoCSV = new System.IO.StreamReader(sPath, System.Text.Encoding.Default);

                    sInsertarOriginal = DllFunciones.GetStringXMLDocument(_oCompany, "eBilling", "eBilling", IDNodoInsert);

                    do
                    {
                        sLinea = ArchivoCSV.ReadLine();

                        if (string.IsNullOrEmpty(sLinea))
                        {
                            DllFunciones.sendMessageBox(_sboapp, "No se encontro el archivo " + NombreArchivoCSV);
                        }
                        else
                        {
                            sArreglo = sLinea.Split(';');

                            if (sArreglo.Length != sCantidadColumnas)
                            {
                                DllFunciones.sendMessageBox(_sboapp, "La cantidad de columnas no es validad para el archivo, " + NombreArchivoCSV);

                            }
                            else
                            {
                                if (NombreArchivoCSV == "Tiposresponsabilidades")
                                {
                                    sCode = Convert.ToString(sArreglo.GetValue(0));
                                    sName = Convert.ToString(sArreglo.GetValue(1));
                                    sDescripcion = Convert.ToString(sArreglo.GetValue(2));

                                    sInsertarProcesado = sInsertarOriginal;

                                    sInsertarProcesado = sInsertarProcesado.Replace("%sCode%", sCode).Replace("%sName%", sName).Replace("%sDescripcion%", sDescripcion);

                                    oInsertar.DoQuery(sInsertarProcesado);
                                }
                                else if (NombreArchivoCSV == "UnidadesdeMedidaDIAN")
                                {
                                    sCode = Convert.ToString(sArreglo.GetValue(0));
                                    sName = Convert.ToString(sArreglo.GetValue(1));

                                    sInsertarProcesado = sInsertarOriginal;

                                    sInsertarProcesado = sInsertarProcesado.Replace("%sCode%", sCode).Replace("%sName%", sName);

                                    oInsertar.DoQuery(sInsertarProcesado);
                                }
                            }
                        }

                        if (ArchivoCSV.EndOfStream)
                        {
                            sSalir = 1;
                            sInsertarOriginal = null;
                            sInsertarProcesado = null;
                            DllFunciones.liberarObjetos(oInsertar);
                        }

                    } while (sSalir == 0);

                    ArchivoCSV.Close();

                    #endregion
                }
                else
                {
                    DllFunciones.liberarObjetos(oValidacion);
                    sVAlidacion = null;
                }
            }
            catch (Exception e)
            {

                sendErrorMessage(_sboapp, e);
            }
        }

        public string SearchCatetoryID(SAPbobsCOM.Company oCompany, string _CategoryName, string DllName)
        {
            string IDCategory;

            sPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            sMotor = Convert.ToString(oCompany.DbServerType);

            sSearchCategory = GetStringXMLDocument(oCompany, DllName, DllName, "GetCategoryIDFormattedSearch");
            sSearchCategory = sSearchCategory.Replace("%CategoryName%", _CategoryName);

            SAPbobsCOM.Recordset oSearchCategory = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            oSearchCategory.DoQuery(sSearchCategory);

            IDCategory = oSearchCategory.Fields.Item(0).Value.ToString();

            liberarObjetos(oSearchCategory);

            return IDCategory;


        }

        /// <summary>
        /// Metodo para generar mensajes de SAP estandar
        /// @achury
        /// </summary>
        /// <param name="message">Mensaje que se desea presentar</param>
        public void sendMessageBox(SAPbouiCOM.Application sboapp, String message)
        {
            try
            {
                sboapp.MessageBox(message, 1);
            }
            catch (Exception ex)
            {
                sendErrorMessage(sboapp, ex);
            }
        }

        /// <summary>
        /// Metodo para generar mensajes de SAP con opcion de Boton
        /// @achury
        /// </summary>
        /// <param name="message">Mensaje que se desea presentar</param>
        /// <param name="btn">Tipo de boton que se desea mostrar 1 OK</param>
        public void sendMessageBox(SAPbouiCOM.Application sboapp, String message, int btn)
        {
            try
            {
                sboapp.MessageBox(message, btn);
            }
            catch (Exception ex)
            {
                sboapp.MessageBox("Incidente: " + ex.Message, 1);
                sboapp.MessageBox("Incidente: " + ex.StackTrace, 1);
            }
        }

        /// <summary>
        /// Metodo para generar mensajes de SAP con respuesta SI o NO
        /// </summary>
        public int sendMessageBoxY_N(SAPbouiCOM.Application sboapp, String message)
        {
            int Respuesta;

            try
            {
                Respuesta = sboapp.MessageBox(message, 1, "Si", "No");

                return Respuesta;

            }
            catch (Exception ex)
            {
                sboapp.MessageBox("Incidente: " + ex.Message, 1);
                sboapp.MessageBox("Incidente: " + ex.StackTrace, 1);

                Respuesta = 2;

                return Respuesta;
            }

        }

        /// <summary>
        /// Metodo para generar mensajes de error con descripción y detalle
        /// @achury
        /// </summary>
        /// <param name="e">Parametro del tipo Exception del error</param>
        public void sendErrorMessage(SAPbouiCOM.Application sboapp, Exception e)
        {
            try
            {
                sboapp.MessageBox("Incidente: " + e.Message, 1);
                sboapp.MessageBox("Incidente: " + e.StackTrace, 1);
            }
            catch (Exception ex)
            {
                sboapp.MessageBox("Incidente: " + ex.Message, 1);
                sboapp.MessageBox("Incidente: " + ex.StackTrace, 1);
            }
        }

        /// <summary>
        /// Método para enviar mensaje por el statusbar
        /// </summary>
        /// <param name="message">Texto que se mostrará</param>
        /// <param name="tiempo">Tiempo que se tarda, valores validas BoMessageTime.bmt_Short bmt.Medium  bmt_Long</param>
        /// <param name="isError">indica si es error o no False o True</param>
        public void sendStatusBarMsg(SAPbouiCOM.Application sboapp, String message, SAPbouiCOM.BoMessageTime tiempo, Boolean isError)
        {
            try
            {
                sboapp.SetStatusBarMessage(message, tiempo, isError);
            }
            catch (Exception ex)
            {
                sboapp.MessageBox("Incidente: " + ex.Message, 1);
                sboapp.MessageBox("Incidente: " + ex.StackTrace, 1);
            }
        }

        public void SelectRowMatrix(SAPbouiCOM.Matrix _oMatrix, ItemEvent _pVal)
        {
            if (_pVal.Row == 0)
            {

            }
            else
            {
                _oMatrix.SelectionMode = BoMatrixSelect.ms_Auto;
                _oMatrix.SelectRow(_pVal.Row, true, true);
            }
        }

        /// <summary>
        /// Método para enviar mensaje de proceso exitoso
        /// </summary>
        public void StatusBar(SAPbouiCOM.Application sboapp, SAPbouiCOM.BoStatusBarMessageType BarType, string Mensaje)
        {
            SAPbouiCOM.BoStatusBarMessageType oStatusBar;

            oStatusBar = BarType;
            sboapp.StatusBar.SetText(Mensaje, SAPbouiCOM.BoMessageTime.bmt_Medium, (SAPbouiCOM.BoStatusBarMessageType)oStatusBar);
        }

        public int FieldSecuence(SAPbobsCOM.Company oCompany, SAPbouiCOM.Application oAplication, string TableName, string FieldId)
        {
            try
            {
                string Query;
                string Secuence;
                SAPbobsCOM.Recordset oRs = null;
                string SiNull = "";
                string DataBase = "";

                oRs = ((SAPbobsCOM.Recordset)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                if (oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    SiNull = "IFNULL";
                    DataBase = oCompany.CompanyDB + ".";
                }
                else
                {
                    SiNull = "ISNULL";
                    DataBase = "";
                }
                Query = "SELECT  " + SiNull + "(MAX(\"FieldID\"), -1) as \"Nro\" FROM " + DataBase + "\"CUFD\" WHERE \"TableID\" = '" + TableName + "' AND \"AliasID\" = '" + FieldId + "'";

                oRs.DoQuery(Query);
                oRs.MoveLast();
                oRs.MoveFirst();
                Secuence = Convert.ToString(oRs.Fields.Item("Nro").Value);
                if (Secuence == "-1")
                {
                    Secuence = Convert.ToString(GetLastSecuence(oCompany, oAplication, TableName) + 1);
                }
                liberarObjetos(oRs);
                return int.Parse(Secuence);
            }
            catch (Exception)
            {
                return -1;
            }

        }

        /// <summary>
        /// Método para obtener la ultima secuencia de la tabla
        /// </summary>
        /// <param name="oCompany"></param>
        /// <param name="oAplication"></param>
        /// <param name="TableName"></param>
        /// <returns></returns>
        public int GetLastSecuence(SAPbobsCOM.Company oCompany, SAPbouiCOM.Application oAplication, string TableName)
        {
            try
            {
                string Query;
                string Secuence;
                SAPbobsCOM.Recordset oRs = null;
                string SiNull = "";
                string DataBase = "";
                oRs = ((SAPbobsCOM.Recordset)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                if (oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    SiNull = "IFNULL";
                    DataBase = oCompany.CompanyDB + ".";
                }
                else
                {
                    SiNull = "ISNULL";
                    DataBase = "";
                }
                Query = "SELECT " + SiNull + "(MAX(\"FieldID\"), 0) AS \"Nro\" FROM  " + DataBase + "CUFD WHERE \"TableID\" = '" + TableName + "'";
                oRs.DoQuery(Query);
                oRs.MoveLast();
                oRs.MoveFirst();
                Secuence = Convert.ToString(oRs.Fields.Item("Nro").Value);
                liberarObjetos(oRs);
                return int.Parse(Secuence);
            }
            catch (Exception)
            {
                return -1;
            }
        }

        /// <summary>
        /// Método para obtener un valor especifico de una fila y columna de un Grid
        /// </summary>
        public string GetFieldGridFromSelectedRow(Grid grid, string columnName)
        {

            if (grid.Rows.SelectedRows.Count == 0) return string.Empty;

            int rowIndex = grid.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder);

            return Convert.ToString(grid.DataTable.GetValue(columnName, rowIndex));

        }

        public int GetFormmatedSearchKey(string _FormID, string _ItemID, SAPbobsCOM.Company oCompany, SAPbouiCOM.Application sboapp)
        {
            int ID = 0;
            string sGetStringXMLDocument = null;

            sGetStringXMLDocument = GetStringXMLDocument(oCompany, "eBilling", "eBilling", "GetFormmatedSearchKey");
            sGetStringXMLDocument = sGetStringXMLDocument.Replace("%FormID%", _FormID).Replace("%ItemID%", _ItemID);

            SAPbobsCOM.Recordset oGetStringXMLDocument = ((SAPbobsCOM.Recordset)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));

            oGetStringXMLDocument.DoQuery(sGetStringXMLDocument);

            ID = Convert.ToInt32(oGetStringXMLDocument.Fields.Item(0).Value.ToString());

            return ID;

        }

        /// <summary>
        /// Método que registra un nuevo AddIn en las tablas del AddOn Basis One
        /// </summary>
        public void InsertAddIn(SAPbouiCOM.Application sboapp, SAPbobsCOM.Company oCompany, string AddIn, string AddInDesc, string Version, string _sNameDB)
        {
            int RsAdd = 0;
            string CurrentUser;

            try
            {
                SAPbobsCOM.UserTable TableUDT;

                CurrentUser = oCompany.UserName;

                TableUDT = oCompany.UserTables.Item("BOAdminAddOn");
                TableUDT.Code = AddIn;
                TableUDT.Name = AddInDesc;
                TableUDT.UserFields.Fields.Item("U_Version").Value = Version;
                TableUDT.UserFields.Fields.Item("U_Status").Value = "I";
                RsAdd = TableUDT.Add();

                if (RsAdd == 0)
                {

                }
                else
                {
                    sboapp.MessageBox(oCompany.GetLastErrorDescription());
                }
            }
            catch (Exception e)
            {
                sendErrorMessage(sboapp, e);
            };
        }

        /// <summary>
        /// Método que registra un nuevo AddIn en las tablas del AddOn Basis One
        /// </summary>
        public void UpdateAddIn(SAPbouiCOM.Application sboapp, SAPbobsCOM.Company oCompany, string AddIn, string Version)
        {
            int RsAdd = 0;

            try
            {
                SAPbobsCOM.UserTable TableUDT;

                TableUDT = oCompany.UserTables.Item("BOAdminAddOn");
                TableUDT.GetByKey(AddIn);
                TableUDT.UserFields.Fields.Item("U_Version").Value = Version;

                RsAdd = TableUDT.Update();

                if (RsAdd == 0)
                {

                }
                else
                {
                    sboapp.MessageBox(oCompany.GetLastErrorDescription());
                }
            }
            catch (Exception e)
            {
                sendErrorMessage(sboapp, e);
            };
        }

        /// <summary>
        /// Método para liberar los objetos
        /// </summary>
        /// <param name="obj">Parametro con el objeto que se requiere liberar</param>
        public void liberarObjetos(object obj)
        {
            if (obj != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
                GC.Collect();
            }

        }

        /// <summary>
        /// Metodo para cargar formulario XML metodo estandar de SAP
        /// @Marlon Gonzalez
        /// </summary>
        public void LoadFromXML(SAPbouiCOM.Application sboapp, string NombreDll, ref string FileName)
        {

            System.Xml.XmlDocument oXmlDoc = null;

            oXmlDoc = new System.Xml.XmlDocument();

            // load the content of the XML File
            string sPath = null;

            sPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

            oXmlDoc.Load(sPath + "\\" + NombreDll + "\\Formularios\\" + FileName);

            // load the form to the SBO application in one batch
            string sXML = oXmlDoc.InnerXml.ToString();
            sboapp.LoadBatchActions(ref sXML);

        }

        /// <summary>
        /// Método que retorna el ID de la tabla OUSR según el código de usuario
        /// </summary>
        /// <param name="oCompany"></param>
        /// <param name="oAplication"></param>
        /// <param name="usrName">Código del Usuario ejemplo(manager)</param>
        /// <returns></returns>
        public String GetUsrID(SAPbobsCOM.Company oCompany, SAPbouiCOM.Application oAplication, string usrName)
        {
            try
            {
                //SAPbobsCOM.Recordset oRs = null;
                string SiNull = "";
                string DataBase = "";
                string Query;
                string usrID = "";

                if (oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    SiNull = "IFNULL";
                    DataBase = ControlChars.Quote + oCompany.CompanyDB + ControlChars.Quote + ".";
                }
                else
                {
                    SiNull = "ISNULL";
                    DataBase = "";
                }

                Query = "SELECT \"USERID\" from BASEDATOS.OUSR where \"USER_CODE\"= '" + usrName + "'";
                Query = Query.Replace("BASEDATOS.", DataBase);
                usrID = GetOneRecordValue(oCompany, oAplication, Query, "USERID");

                if (usrID == null)
                {
                    throw new Exception($"Usuario '{usrName}' no encontrado");
                }
                return usrID;

            }
            catch (Exception ex)
            {
                //Console.WriteLine($"Error en la consulta de usuario: {ex.Message}");
                sendErrorMessage(oAplication, ex);
                return null;
            }

        }

        /// <summary>
        /// Método que retorna el Nombre de usuario de la tabla OUSR según el código de usuario
        /// </summary>
        /// <param name="oCompany"></param>
        /// <param name="oAplication"></param>
        /// <param name="usrCode">Código del Usuario ejemplo(manager)</param>
        /// <returns></returns>
        public String GetUsrName(SAPbobsCOM.Company oCompany, SAPbouiCOM.Application oAplication, string usrCode)
        {
            try
            {
                SAPbobsCOM.Recordset oRs = null;
                string SiNull = "";
                string DataBase = "";
                string Query;
                string usrID = "";

                if (oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    SiNull = "IFNULL";
                    DataBase = ControlChars.Quote + oCompany.CompanyDB + ControlChars.Quote + ".";
                }
                else
                {
                    SiNull = "ISNULL";
                    DataBase = "";
                }

                Query = "SELECT \"U_NAME\" from BASEDATOS.OUSR where \"USER_CODE\"= '" + usrCode + "'";
                Query = Query.Replace("BASEDATOS.", DataBase);
                usrID = GetOneRecordValue(oCompany, oAplication, Query, "U_NAME");

                if (usrID == null)
                {
                    throw new Exception($"Usuario '{usrCode}' no encontrado");
                }
                return usrID;

            }
            catch (Exception ex)
            {
                //Console.WriteLine($"Error en la consulta de usuario: {ex.Message}");
                sendErrorMessage(oAplication, ex);
                return null;
            }

        }

        /// <summary>
        /// Método para ejecutar un query y retornar el recordset con el resultado 
        /// </summary>
        /// <param name="oCompany"></param>
        /// <param name="oAplication"></param>
        /// <param name="Query">Query que se requiere ejecutar</param>
        /// <returns></returns>
        public SAPbobsCOM.Recordset ExecRecordSet(SAPbobsCOM.Company oCompany, SAPbouiCOM.Application oAplication, string Query)
        {
            SAPbobsCOM.Recordset oRs = null;
            try
            {
                oRs = ((SAPbobsCOM.Recordset)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                oRs.DoQuery(Query);
                return oRs;
            }
            catch (Exception ex)
            {
                sendErrorMessage(oAplication, ex);
                return null;
            }
        }

        /// <summary>
        /// Método para ejecutar un query y retornar el recordset con el resultado con el valor del campo específico
        /// </summary>
        /// <param name="oCompany"></param>
        /// <param name="oAplication"></param>
        /// <param name="Query">Query de la consulta de los datos</param>
        /// <param name="oField">Campo solicitado</param>
        /// <returns></returns>
        public string GetOneRecordValue(SAPbobsCOM.Company oCompany, SAPbouiCOM.Application oAplication, string Query, string oField)
        {
            SAPbobsCOM.Recordset oRs = null;
            try
            {
                oRs = ExecRecordSet(oCompany, oAplication, Query);
                if (oRs != null)
                {
                    if (oRs.RecordCount > 0)
                    {
                        oRs.MoveFirst();
                        return (oRs.Fields.Item(oField).Value).ToString();
                    }
                    else
                    {
                        return null;
                    }
                }
                else
                {
                    return null;
                }
                liberarObjetos(oRs);
            }
            catch (Exception ex)
            {
                sendErrorMessage(oAplication, ex);
                liberarObjetos(oRs);
                return null;
            }

        }

        /// <summary>
        /// Método para crear una barra de progreso
        /// </summary>
        public void ProgressBar(SAPbobsCOM.Company oCompany, SAPbouiCOM.Application sboapp, int ValorFinal, int ValoraIncrementar, string TextProgressBar)
        {
            Contador = Contador + ValoraIncrementar;

            if (oProgressBar != null)
            {
                //Contador = Contador + ValoraIncrementar;
                oProgressBar.Text = TextProgressBar;
                oProgressBar.Value += ValoraIncrementar;
                if (Contador == ValorFinal)
                {
                    oProgressBar.Stop();
                    liberarObjetos(oProgressBar);
                }
            }
            else
            {
                SAPbouiCOM.Application SBO_Application = sboapp;
                oProgressBar = SBO_Application.StatusBar.CreateProgressBar("AddOnBO", ValorFinal, true);
                oProgressBar.Text = TextProgressBar;
                oProgressBar.Value += Contador;
            }
        }

        /// <summary>
        /// Método que retorna el código string correspondiente, equivalente a funciones de vb.net 
        /// </summary>
        public sealed class ControlChars
        {
            public const char Back = '\b';
            public const char Cr = '\r';
            public const string CrLf = "\r\n";
            public const char FormFeed = '\f';
            public const char Lf = '\n';
            public const string NewLine = "\r\n";
            public const char NullChar = '\0';
            public const char Quote = '"';
            public const char Tab = '\t';
            public const char VerticalTab = '\v';
        }

        /// <summary>
        /// Método que retorna el string de un archivo xml
        /// </summary>
        /// 
        public string GetStringXMLDocument(SAPbobsCOM.Company _oCompany, string DllName, string FileNameXML, string IDNodo)
        {
            string sFileXML = null;


            sPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            sMotor = Convert.ToString(_oCompany.DbServerType);
            sNameDB = _oCompany.CompanyDB;

            XmlDocument xmlDocument;
            xmlDocument = new XmlDocument();

            switch (sMotor)
            {
                case "dst_MSSQL2012":
                case "dst_MSSQL2014":
                case "dst_MSSQL2016":
                case "dst_MSSQL2017":

                    xmlDocument.Load(sPath + "\\" + DllName + "\\Queries\\" + FileNameXML + "SQL.xml");
                    sFileXML = xmlDocument.SelectSingleNode("Queries/" + IDNodo + "").InnerText;


                    break;

                case "dst_HANADB":

                    xmlDocument.Load(sPath + "\\" + DllName + "\\Queries\\" + FileNameXML + "HANA.xml");
                    sFileXML = xmlDocument.SelectSingleNode("Queries/" + IDNodo + "").InnerText;
                    sFileXML = sFileXML.Replace("%sNameDB%", sNameDB);

                    break;

                default:
                    break;
            }
            return sFileXML;
        }

        /// <summary>
        /// Returna el color BLANCO para SAP B1
        /// </summary>
        /// 
        public int ColorSB1_BLANCO()
        {
            int clInt = 255 | (255 << 8) | (255 << 16);
            return clInt;

        }

        /// <summary>
        /// Returna el color PLATEADO para SAP B1
        /// </summary>
        /// 
        public int ColorSB1_PLATEADO()
        {
            int clInt = 192 | (192 << 8) | (192 << 16);
            return clInt;

        }

        /// <summary>
        /// Returna el color GRIS para SAP B1
        /// </summary>
        /// 
        public int ColorSB1_GRIS()
        {
            int clInt = 128 | (128 << 8) | (128 << 16);
            return clInt;

        }

        /// <summary>
        /// Returna el color ROJO para SAP B1
        /// </summary>
        /// 
        public int ColorSB1_ROJO()
        {
            int clInt = 255 | (0 << 8) | (0 << 16);
            return clInt;

        }

        /// <summary>
        /// Returna el color MARRON para SAP B1
        /// </summary>
        /// 
        public int ColorSB1_MARRON()
        {
            int clInt = 128 | (0 << 8) | (0 << 16);
            return clInt;

        }

        /// <summary>
        /// Returna el color AMARILLO para SAP B1
        /// </summary>
        /// 
        public int ColorSB1_AMARILLO()
        {
            int clInt = 255 | (255 << 8) | (0 << 16);
            return clInt;

        }

        /// <summary>
        /// Returna el color LIMA para SAP B1
        /// </summary>
        /// 
        public int ColorSB1_LIMA()
        {
            int clInt = 0 | (255 << 8) | (0 << 16);
            return clInt;

        }

        /// <summary>
        /// Returna el color VERDE para SAP B1
        /// </summary>
        /// 
        public int ColorSB1_VERDE()
        {
            int clInt = 0 | (128 << 8) | (0 << 16);
            return clInt;

        }

        /// <summary>
        /// Returna el color AGUA para SAP B1
        /// </summary>
        /// 
        public int ColorSB1_AGUA()
        {
            int clInt = 0 | (255 << 8) | (255 << 16);
            return clInt;

        }

        /// <summary>
        /// Returna el color VERDE_AZULADO para SAP B1
        /// </summary>
        /// 
        public int ColorSB1_VERDE_AZULADO()
        {
            int clInt = 0 | (128 << 8) | (128 << 16);
            return clInt;

        }

        /// <summary>
        /// Returna el color AZUL para SAP B1
        /// </summary>
        public int ColorSB1_AZUL()
        {
            int clInt = 0 | (0 << 8) | (255 << 16);
            return clInt;

        }

        /// <summary>
        /// Returna el color NAVY para SAP B1
        /// </summary>
        public int ColorSB1_NAVY()
        {
            int clInt = 0 | (0 << 8) | (255 << 16);
            return clInt;

        }

        /// <summary>
        /// Returna el color FUCHSIA para SAP B1
        /// </summary>
        public int ColorSB1_FUCHSIA()
        {
            int clInt = 255 | (0 << 8) | (255 << 16);
            return clInt;

        }

        /// <summary>
        /// Returna el color PURPURA para SAP B1
        /// </summary>
        public int ColorSB1_PURPURA()
        {
            int clInt = 128 | (0 << 8) | (128 << 16);
            return clInt;

        }

        /// <summary>
        /// Returna el color NARANJA para SAP B1
        /// </summary>
        public int ColorSB1_NARANJA()
        {
            int clInt = 255 | (127 << 8) | (0 << 16);
            return clInt;

        }

        public string VersionDll()
        {
            try
            {

                Assembly Assembly = Assembly.LoadFrom("BOFunciones.dll");
                Version vVersion = Assembly.GetName().Version;

                String VersionDll = vVersion.ToString();

                return VersionDll;
            }
            catch (Exception)
            {

                throw;
            }

        }

        public void Logger(String LogMessage, string Path)
        {
            StreamWriter Log = File.AppendText(Path);
            Log.WriteLine("----------------------------------------------------");
            Log.WriteLine("{0} {1}", DateTime.Now, LogMessage);
            Log.Close();
            StreamReader r = File.OpenText(Path);
            DumpLog(r);
        }

        private static void DumpLog(StreamReader r)
        {
            // While not at the end of the file, read and write lines.
            String line;
            while ((line = r.ReadLine()) != null)
            {
                Console.WriteLine(line);
            }
            r.Close();
        }

        public string GetSeparatorMachine()
        {
            string SeparatorMachine = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.CurrencyDecimalSeparator;

            return SeparatorMachine;

        }

        public string GetSeparatorSAP(SAPbobsCOM.Company _oCompany)
        {
            string sSeparador = GetStringXMLDocument(_oCompany, "Core", "ValidacionAddOnBO", "GetSepartorDecimal");

            SAPbobsCOM.Recordset oRSepDec = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            oRSepDec.DoQuery(sSeparador);

            string SeparatorSAP = Convert.ToString(oRSepDec.Fields.Item("DecSep").Value.ToString());

            liberarObjetos(oRSepDec);

            return SeparatorSAP;

        }


    }
}
