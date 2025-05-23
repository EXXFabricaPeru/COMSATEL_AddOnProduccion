using System;
using System.Threading;
using AddonProduccionEnsDes.conexion;
using SAPbouiCOM;

namespace AddonProduccionEnsDes.commons
{
    public class FormCommon
    {
        #region Generic
        private const string RS_VALUE = "Value";
        private const string RS_NAME = "Name";

        public SAPbouiCOM.Form CreateForm(SAPbobsCOM.Company company, SAPbouiCOM.Application application, string resource, string formName)
        {
            SAPbouiCOM.Form mForm = null;

            try
            {
                SAPbouiCOM.FormCreationParams fCreationParams = application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                fCreationParams.XmlData = resource;
                fCreationParams.FormType = formName;
                fCreationParams.UniqueID = formName + DateTime.Now.ToString("hhmmss");
                mForm = application.Forms.AddEx(fCreationParams);
                mForm.Visible = false;
            }
            catch (Exception ex)
            {
                StatusMessageError("Error creando formulario " + formName + ". Excepción :" + ex.Message);
            }

            return mForm;
        }

        public static void StatusMessageError(string message)
        {
            Conexion.application.StatusBar.SetText(Constants.PREFIX_MSG_ADDON + message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        }
        public static void StatusMessageInfo(string message)
        {
            Conexion.application.StatusBar.SetText(Constants.PREFIX_MSG_ADDON + message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None);
        }
        public static void StatusMessageWarning(string message)
        {
            Conexion.application.StatusBar.SetText(Constants.PREFIX_MSG_ADDON + message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
        }
        public static void StatusMessageSuccess(string message)
        {
            Conexion.application.StatusBar.SetText(Constants.PREFIX_MSG_ADDON + message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
        }

        internal static void LiberarObjetoGenerico(Object objeto)
        {
            try
            {
                if (objeto != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objeto);

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(Constants.PREFIX_MSG_ADDON + " Error Liberando Objeto: " + ex.Message);
            }
        }
        internal static void SumColumna(SAPbouiCOM.Column oColumn)
        {
            oColumn.ColumnSetting.SumType = BoColumnSumType.bst_Auto;
        }
        internal static void InstanciateCombo(SAPbouiCOM.ComboBox ComboBox, string Query = "", bool onlyDesc = true)
        {
            SAPbobsCOM.Recordset ComboRS = (SAPbobsCOM.Recordset)Conexion.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            if (!string.IsNullOrEmpty(Query))
            {
                while (ComboBox.ValidValues.Count != 0)
                {
                    ComboBox.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                }
                ComboRS.DoQuery(Query);
                try
                {
                    while (!ComboRS.EoF)
                    {
                        ComboBox.ValidValues.Add((string)ComboRS.Fields.Item(RS_VALUE).Value.ToString(), (string)ComboRS.Fields.Item(RS_NAME).Value.ToString());
                        ComboRS.MoveNext();
                    }
                }
                catch (Exception)
                {

                }
                ComboBox.Item.Enabled = true;
            }
            if (onlyDesc) ComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;


            System.Runtime.InteropServices.Marshal.ReleaseComObject(ComboRS);
        }
        internal static void InstanciateComboColumn(SAPbouiCOM.ComboBoxColumn ComboBox, string Query)
        {
            SAPbobsCOM.Recordset ComboRS = (SAPbobsCOM.Recordset)Conexion.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            while (ComboBox.ValidValues.Count != 0)
            {
                ComboBox.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
            }
            ComboRS.DoQuery(Query);
            while (!ComboRS.EoF)
            {
                ComboBox.ValidValues.Add((string)ComboRS.Fields.Item(RS_VALUE).Value.ToString(), (string)ComboRS.Fields.Item(RS_NAME).Value.ToString());
                ComboRS.MoveNext();
            }
            //ComboBox.Item.Enabled = true;
            ComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(ComboRS);
        }
        internal bool OpenFileDialog(SAPbouiCOM.Form oForm, string FILE_TXT)
        {
            GetFileNameClass oGetFileName = new GetFileNameClass();
            oGetFileName.Filter = "Archivo excel (.xlsx)|*.xlsx";
            Thread threadGetFile = new Thread(new ThreadStart(oGetFileName.GetFileName));
            threadGetFile.SetApartmentState(ApartmentState.STA);
            try
            {
                threadGetFile.Start();
                while (!threadGetFile.IsAlive) ; // Wait for thread to get started
                Thread.Sleep(1);  // Wait a sec more
                threadGetFile.Join();    // Wait for thread to end
                SAPbouiCOM.EditText txRuta = (SAPbouiCOM.EditText)oForm.Items.Item(FILE_TXT).Specific;
                txRuta.Value = oGetFileName.FileName;

            }
            catch (Exception ex)
            {
                StatusMessageWarning(string.Format("openFileDialog():{0}", ex.Message));
            }
            finally
            {
                threadGetFile = null;
                oGetFileName = null;
            }
            return true;
        }
        internal static string GenericQuery(string query)
        {
            SAPbobsCOM.Recordset oRS = Conexion.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string valor = "";
            try
            {
                oRS.DoQuery(query);
                if (oRS.RecordCount > 0)
                {
                    valor = oRS.Fields.Item("Value").Value.ToString().Trim();
                }
            }
            catch (Exception ex)
            {

            }
            finally { LiberarObjetoGenerico(oRS); }

            return valor;
        }

        internal static string GetValueComboRuta(string table, string field, string code)
        {
            SAPbobsCOM.Recordset oRS = null;
            try
            {

                oRS = Conexion.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRS.DoQuery(Queries.GetRutaValues(code));
                return oRS.Fields.Item("Value").Value.ToString();
            }
            catch (Exception)
            {
                return "";
            }
            finally
            {
                LiberarObjetoGenerico(oRS);
            }
        }

        internal static string GetConfig(string id, DateTime? fecha = null)
        {
            SAPbobsCOM.Recordset oRS = null;
            try
            {
                switch (id)
                {
                    case "RATIO":
                        oRS = Conexion.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        //oRS.DoQuery(Queries.GetRatio(((DateTime)fecha).ToString(Constantes.PERDATE_FORMAT)));
                        if (oRS.RecordCount == 0)
                            StatusMessageWarning(string.Format("GetConfig():{0}", "Colocar tasa en periodo."));

                        return oRS.Fields.Item("Value").Value.ToString();
                    default:
                        oRS = Conexion.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        oRS.DoQuery(Queries.GetCFGValue(id));
                        return oRS.Fields.Item("Value").Value.ToString();
                }
            }
            catch (Exception)
            {
                return "";
            }
            finally
            {
                LiberarObjetoGenerico(oRS);
            }
        }

        internal static string FormatPrice(string price)
        {
            return price.Replace("PEN", "").Replace(Constants.MAIN_CURR, "").Replace("USD", "").Replace("S/", "");
        }

        internal static bool CheckShowAux(string id)
        {
            SAPbobsCOM.Recordset oRS = null;
            try
            {
                oRS = Conexion.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRS.DoQuery(Queries.GetCheckCFGAux(id));
                return oRS.Fields.Item("Check").Value.ToString().Equals("Y");
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                LiberarObjetoGenerico(oRS);
            }
        }
        #endregion

    }
}
