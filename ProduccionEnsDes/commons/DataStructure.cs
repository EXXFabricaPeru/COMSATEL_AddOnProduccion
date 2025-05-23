using System;
using System.Deployment.Application;
using System.Reflection;
using AddonProduccionEnsDes.conexion;
using AddonProduccionEnsDes.data_schema;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AddonProduccionEnsDes.commons
{
    public class DataStructure : FormCommon
    {
        #region _Attributes_

        int m_iErrCode = 0;
        string m_sErrMsg = "";
        private string m_sNombreAddon = Properties.Resources.NombreAddon;
        #endregion

        #region _Constructor_

        public DataStructure()
        {
            try
            {
                StatusMessageInfo("Verificando estructura de datos del AddOn...");
                if (ValidaVersion(m_sNombreAddon, m_sVersion))
                {
                    RegistrarVersion(m_sNombreAddon, m_sVersion);
                    CrearTablasADDON();
                    CrearCamposADDON();
                    CrearObjetosADDON();
                }
                StatusMessageInfo("Verificación completada.");
            }
            catch (Exception ex)
            {
                StatusMessageError(Properties.Resources.NombreAddon + " Error: EstructuraDatos.cs > EstructuraDatos():" + ex.Message);
            }
        }

        #endregion

        #region _Methods_

        private void CrearTablasADDON()
        {
            StatusMessageWarning("Verificando Tablas...");
            foreach (var item in SchemaAddon.tablesGeneric())
            {
                CreaTablaMD(item.Key, item.Value, SAPbobsCOM.BoUTBTableType.bott_NoObjectAutoIncrement);
            }
            foreach (var item in SchemaAddon.tablesGenericNoAuto())
            {
                CreaTablaMD(item.Key, item.Value, SAPbobsCOM.BoUTBTableType.bott_NoObject);
            }
            foreach (var item in SchemaAddon.tablesMasterH())
            {
                CreaTablaMD(item.Key, item.Value, SAPbobsCOM.BoUTBTableType.bott_MasterData);
            }
            foreach (var item in SchemaAddon.tablesMasterD())
            {
                CreaTablaMD(item.Key, item.Value, SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
            }
            foreach (var item in SchemaAddon.tablesDocsH())
            {
                CreaTablaMD(item.Key, item.Value, SAPbobsCOM.BoUTBTableType.bott_Document);
            }
            foreach (var item in SchemaAddon.tablesDocsD())
            {
                CreaTablaMD(item.Key, item.Value, SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
            }
        }

        private void CrearCamposADDON()
        {
            StatusMessageWarning("Verificando campos...");
            foreach (var item in SchemaAddon.camposTB())
            {
                CreaCampoMD(item.nombre_tabla, item.nombre_campo, item.descrp_campo,
                            item.tipo_campo, item.subtipo_campo, item.tamano, item.obligatorio,
                            item.validValues, item.validDescription, item.valorPorDef, item.tablaVinculada, item.objetoVinculado);
            }
        }

        private void CrearObjetosADDON()
        {
            StatusMessageWarning("Verificando Objetos...");

            foreach (var item in SchemaAddon.objetosADDON())
            {
                CreaUDOMD(item.code, item.name, item.tableName, item.findColumns, item.childTables,
                          item.canCancel, item.canClose, item.canDelete, item.canCreateDefaultForm,
                          item.formColumns, item.canFind, item.canLog, item.objectType, item.manageSeries,
                          item.enableEnhancedForm, item.rebuildEnhancedForm, item.childFormColumns);
            }
        }

        #endregion

        #region _Functions_

        private bool ValidaVersion(string NombreAddon, string Version)
        {
            bool bRetorno = false;
            SAPbobsCOM.UserTable oUT = null;
            SAPbobsCOM.Recordset oRS = null;
            string NombreTabla = "";
            try
            {
                NombreTabla = NombreAddon.ToUpper();
                try
                {
                    oUT = Conexion.company.UserTables.Item(NombreTabla);
                }
                catch (Exception ex)
                {
                    if (ex.Message.ToLower().Contains("invalid field name")) oUT = null;
                    else throw ex;
                }
                if (oUT == null)
                {
                    CreaTablaMD(NombreTabla, "", SAPbobsCOM.BoUTBTableType.bott_NoObject);
                    bRetorno = true;
                }
                else
                {
                    oRS = (SAPbobsCOM.Recordset)Conexion.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRS.DoQuery(Queries.ConsultaTablaConfiguracion(Conexion.company.DbServerType, NombreAddon, "", true));
                    if (oRS.RecordCount == 0)
                    {
                        bRetorno = true;
                        Conexion.application.StatusBar.SetText(Properties.Resources.NombreAddon + " Actualizará la estructura de datos",
                            SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    }
                    else
                    {
                        var a = int.Parse(Version.Replace(".", "").ToString());
                        var b = int.Parse(oRS.Fields.Item("code").Value.ToString().Replace(".", ""));
                        if (int.Parse(Version.Replace(".", "").ToString()) > int.Parse(oRS.Fields.Item("code").Value.ToString().Replace(".", "")))
                        {
                            bRetorno = true;
                            Conexion.application.StatusBar.SetText(Properties.Resources.NombreAddon + " Actualizará la estructura de datos",
                            SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        }

                        if (int.Parse(Version.Replace(".", "").ToString()) < int.Parse(oRS.Fields.Item("code").Value.ToString().Replace(".", "")))
                            Conexion.application.StatusBar.SetText(Properties.Resources.NombreAddon + " Detectó una version superior para este Addon",
                            SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                }
            }
            catch (Exception ex)
            {
                Conexion.application.StatusBar.SetText(Properties.Resources.NombreAddon +
                    " Error: EstructuraDatos.cs > ValidaVersion():" + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                LiberarObjetoGenerico(oUT);
                LiberarObjetoGenerico(oRS);
                oRS = null;
                oUT = null;
            }
            return bRetorno;
        }

        private void RegistrarVersion(string NombreAddon, string Version)
        {
            SAPbobsCOM.UserTable oUT = null;
            string NombreTabla = "";
            try
            {
                NombreTabla = NombreAddon.ToUpper();
                oUT = Conexion.company.UserTables.Item(NombreTabla);
                oUT.Code = Version;
                oUT.Name = NombreAddon + " V-" + Version;
                m_iErrCode = oUT.Add();
                if (m_iErrCode == 0)
                    StatusMessageSuccess(Properties.Resources.NombreAddon + " Se ingreso un nuevo registro a la BD ");
                else
                    StatusMessageError(Properties.Resources.NombreAddon + " Error ingresar el registro en la tabla: " + NombreTabla);
            }
            catch (Exception ex)
            {
                StatusMessageError(Properties.Resources.NombreAddon + " Error: EstructuraDatos.cs > RegistrarVersion():" + ex.Message);
            }
            finally
            {
                LiberarObjetoGenerico(oUT);
                oUT = null;
            }
        }

        private bool CreaTablaMD(string NombTabla, string DescTabla, SAPbobsCOM.BoUTBTableType tipoTabla)
        {
            SAPbobsCOM.UserTablesMD oUserTablesMD = null;
            try
            {
                oUserTablesMD = null;
                oUserTablesMD = (SAPbobsCOM.UserTablesMD)Conexion.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
                if (!oUserTablesMD.GetByKey(NombTabla))
                {
                    oUserTablesMD.TableName = NombTabla;
                    oUserTablesMD.TableDescription = DescTabla;
                    oUserTablesMD.TableType = tipoTabla;

                    m_iErrCode = oUserTablesMD.Add();
                    if (m_iErrCode != 0)
                    {
                        Conexion.company.GetLastError(out m_iErrCode, out m_sErrMsg);
                        StatusMessageError(Properties.Resources.NombreAddon + " Error al crear  tabla: " + NombTabla);
                        return false;
                    }
                    else
                        StatusMessageSuccess(Properties.Resources.NombreAddon + " Se ha creado la tabla: " + NombTabla);

                    LiberarObjetoGenerico(oUserTablesMD);
                    oUserTablesMD = null;
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                StatusMessageError(Properties.Resources.NombreAddon + " Error: EstructuraDatos.cs > CreaTablaMD():" + ex.Message);
                return false;
            }
            finally
            {
                LiberarObjetoGenerico(oUserTablesMD);
                oUserTablesMD = null;
            }
        }

        private void CreaCampoMD(string NombreTabla, string NombreCampo, string DescCampo, SAPbobsCOM.BoFieldTypes TipoCampo, SAPbobsCOM.BoFldSubTypes SubTipo, int Tamano, SAPbobsCOM.BoYesNoEnum Obligatorio, string[] validValues, string[] validDescription, string valorPorDef, string tablaVinculada, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum objetoVinculado)
        {
            SAPbobsCOM.UserFieldsMD oUserFieldsMD = null;
            try
            {
                int existeCampo = 0;

                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Conexion.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string cadena = "select \"FieldID\" from CUFD where (\"TableID\"='" + NombreTabla + "' or \"TableID\"='@" + NombreTabla + "') and \"AliasID\"='" + NombreCampo + "'";
                rs.DoQuery(cadena);

                existeCampo = rs.RecordCount;
                int FieldID = Convert.ToInt32(rs.Fields.Item(0).Value);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                rs = null;

                if (existeCampo == 0)
                {
                    if (NombreTabla == null) NombreTabla = "";
                    if (NombreCampo == null) NombreCampo = "";
                    if (Tamano == 0) Tamano = 10;
                    if (validValues == null) validValues = new string[0];
                    if (validDescription == null) validDescription = new string[0];
                    if (valorPorDef == null) valorPorDef = "";
                    if (tablaVinculada == null) tablaVinculada = "";
                    oUserFieldsMD = (SAPbobsCOM.UserFieldsMD)Conexion.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                    oUserFieldsMD.TableName = NombreTabla;
                    oUserFieldsMD.Name = NombreCampo;
                    oUserFieldsMD.Description = DescCampo;
                    oUserFieldsMD.Type = TipoCampo;
                    if (TipoCampo != SAPbobsCOM.BoFieldTypes.db_Date) oUserFieldsMD.EditSize = Tamano;
                    oUserFieldsMD.SubType = SubTipo;

                    if (tablaVinculada != "") oUserFieldsMD.LinkedTable = tablaVinculada;
                    if (objetoVinculado != SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone) oUserFieldsMD.LinkedSystemObject = objetoVinculado;
                    if (validValues.Length > 0)
                    {
                        for (int i = 0; i <= (validValues.Length - 1); i++)
                        {
                            oUserFieldsMD.ValidValues.Value = validValues[i];
                            if (validDescription.Length > 0) oUserFieldsMD.ValidValues.Description = validDescription[i];
                            else oUserFieldsMD.ValidValues.Description = validValues[i];
                            oUserFieldsMD.ValidValues.Add();
                        }
                    }
                    oUserFieldsMD.Mandatory = Obligatorio;
                    if (valorPorDef != "") oUserFieldsMD.DefaultValue = valorPorDef;

                    m_iErrCode = oUserFieldsMD.Add();
                    if (m_iErrCode != 0)
                    {
                        Conexion.company.GetLastError(out m_iErrCode, out m_sErrMsg);
                        if ((m_iErrCode != -5002) && (m_iErrCode != -2035))
                            Conexion.application.StatusBar.SetText(Properties.Resources.NombreAddon + " Error al crear campo de usuario: " + NombreCampo
                                + "en la tabla: " + NombreTabla + " Error: " + m_sErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                    else
                        Conexion.application.StatusBar.SetText(Properties.Resources.NombreAddon + " Se ha creado el campo de usuario: " + NombreCampo
                                + " en la tabla: " + NombreTabla, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
            }
            catch (Exception ex)
            {
                Conexion.application.StatusBar.SetText(Properties.Resources.NombreAddon +
                    " Error: EstructuraDatos.cs > CreaCampoMD():" + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                LiberarObjetoGenerico(oUserFieldsMD);
                oUserFieldsMD = null;
            }
        }

        private bool CreaUDOMD(string sCode, string sName, string sTableName, string[] sFindColumns,
            string[] sChildTables, SAPbobsCOM.BoYesNoEnum eCanCancel, SAPbobsCOM.BoYesNoEnum eCanClose,
            SAPbobsCOM.BoYesNoEnum eCanDelete, SAPbobsCOM.BoYesNoEnum eCanCreateDefaultForm, string[] sFormColumns,
            SAPbobsCOM.BoYesNoEnum eCanFind, SAPbobsCOM.BoYesNoEnum eCanLog, SAPbobsCOM.BoUDOObjType eObjectType,
            SAPbobsCOM.BoYesNoEnum eManageSeries, SAPbobsCOM.BoYesNoEnum eEnableEnhancedForm,
            SAPbobsCOM.BoYesNoEnum eRebuildEnhancedForm, string[] sChildFormColumns)
        {
            SAPbobsCOM.UserObjectsMD oUserObjectMD = null;
            int i_Result = 0;
            string s_Result = "";

            try
            {
                oUserObjectMD = (SAPbobsCOM.UserObjectsMD)Conexion.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);

                if (!oUserObjectMD.GetByKey(sCode))
                {
                    oUserObjectMD.Code = sCode;
                    oUserObjectMD.Name = sName;
                    oUserObjectMD.ObjectType = eObjectType;
                    oUserObjectMD.TableName = sTableName;
                    oUserObjectMD.CanCancel = eCanCancel;
                    oUserObjectMD.CanClose = eCanClose;
                    oUserObjectMD.CanDelete = eCanDelete;
                    oUserObjectMD.CanCreateDefaultForm = eCanCreateDefaultForm;
                    oUserObjectMD.EnableEnhancedForm = eEnableEnhancedForm;
                    oUserObjectMD.RebuildEnhancedForm = eRebuildEnhancedForm;
                    oUserObjectMD.CanFind = eCanFind;
                    oUserObjectMD.CanLog = eCanLog;
                    oUserObjectMD.ManageSeries = eManageSeries;

                    if (sFindColumns != null)
                    {
                        for (int i = 0; i < sFindColumns.Length; i++)
                        {
                            oUserObjectMD.FindColumns.ColumnAlias = sFindColumns[i];
                            oUserObjectMD.FindColumns.Add();
                        }
                    }
                    if (sChildTables != null)
                    {
                        for (int i = 0; i < sChildTables.Length; i++)
                        {
                            oUserObjectMD.ChildTables.TableName = sChildTables[i];
                            oUserObjectMD.ChildTables.Add();
                        }
                    }
                    if (sFormColumns != null)
                    {
                        oUserObjectMD.UseUniqueFormType = SAPbobsCOM.BoYesNoEnum.tYES;

                        for (int i = 0; i < sFormColumns.Length; i++)
                        {
                            oUserObjectMD.FormColumns.FormColumnAlias = sFormColumns[i];
                            oUserObjectMD.FormColumns.Add();
                        }
                    }
                    if (sChildFormColumns != null)
                    {
                        if (sChildTables != null)
                        {
                            for (int i = 0; i < sChildFormColumns.Length; i++)
                            {
                                oUserObjectMD.FormColumns.SonNumber = 1;
                                oUserObjectMD.FormColumns.FormColumnAlias = sChildFormColumns[i];
                                oUserObjectMD.FormColumns.Add();
                            }
                        }
                    }

                    i_Result = oUserObjectMD.Add();

                    if (i_Result != 0)
                    {
                        Conexion.company.GetLastError(out i_Result, out s_Result);
                        Conexion.application.StatusBar.SetText(Properties.Resources.NombreAddon +
                            " Error: EstructuraDatos.cs > CreaUDOMD(): " + s_Result + ", creando el UDO " + sCode + ".",
                            SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        return false;
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                Conexion.application.StatusBar.SetText(Properties.Resources.NombreAddon +
                    " Error: EstructuraDatos.cs > CreaUDOMD():" + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }
            finally
            {
                LiberarObjetoGenerico(oUserObjectMD);
            }
        }

        public string m_sVersion
        {
            get
            {
                return ApplicationDeployment.IsNetworkDeployed
                       ? ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString()
                       : Assembly.GetExecutingAssembly().GetName().Version.ToString();
            }
        }
        #endregion

    }
}
