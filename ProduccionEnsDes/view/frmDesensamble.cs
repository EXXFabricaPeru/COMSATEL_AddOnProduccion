using AddonProduccionEnsDes.commons;
using AddonProduccionEnsDes.conexion;
using AddonProduccionEnsDes.data_schema;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using SAPbobsCOM;
using SAPbouiCOM;
using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace AddonProduccionEnsDes.view
{
    public class frmDesensamble : FormCommon, commons.IForm
    {
        #region variables
        private SAPbouiCOM.Form mForm;
        private SAPbouiCOM.DBDataSource dsHEAD;
        private SAPbouiCOM.DBDataSource dsDETA;
        private SAPbouiCOM.UserDataSource udRUTA;
        private SAPbouiCOM.Matrix oMatrix;
        private SAPbouiCOM.Item oEdtDate, oEdtWhs;
        private SAPbouiCOM.Item btnOW, btnNS, btnActS, btnAdd, btnCE;

        //CONST PARA LAYOUT
        public const string TYPENAME = "Formulacion";
        public const string ADDONNAME = "ListaMateriales";

        private const string EDT_DATE = "edtDate", EDT_WHS = "edtWHS"; //EditTexts Porc
        private const string BTN_OK = "1", BTN_ORDFAB = "btnOrd", BTN_ENSAMBLAR = "btnExe", BTN_ADD = "btnAdd", BTN_SERIE = "btnSerie", BTN_FILE = "btnFile", BTN_ACTSERIE = "btnAct";//Buttons
        private const string CBO_RUTA = "cboRuta";//Combo
        private const string MTX_MAIN = "mtxMain"; //Matrix
        private const string UD_RUTA = "UdRuta";
        private const string C_CODRUT = "Col_0", C_CODRUTO = "Col_2";


        string codeLM = "1", origRuta = "";

        //Right Click
        private string ItemUIDRightClick;
        private int RowItemRightClick;
        private object eCommon;
        #endregion

        public frmDesensamble(Dictionary<string, commons.IForm> dictionary)
        {
            try
            {
                mForm = CreateForm(Conexion.company, Conexion.application, Properties.Resources.frmDesensamble2, FormName.DESENSAMBLE);
                if (mForm != null)
                {
                    mForm.Freeze(true);
                    dictionary.Add(getFormUID(), (commons.IForm)this);
                    Initializer();
                    mForm.Visible = true;
                }
                else
                    StatusMessageError("Construct: No se pudo crear el formulario " + FormName.FORMUL_RLM);
            }
            catch (Exception)
            {
                throw;
            }
            finally { if (mForm != null) mForm.Freeze(false); }
        }


        #region _EVENTOS_ITEMEVENT

        //Principal

        private void Initializer()
        {
            try
            {
                mForm.Freeze(true);
                if (dsHEAD == null) dsHEAD = mForm.DataSources.DBDataSources.Item($"@{SCDesensamble.TABLE_CABE}");
                dsDETA = mForm.DataSources.DBDataSources.Item($"@{SCDesensamble.TABLE_DET1}");
                oMatrix = (SAPbouiCOM.Matrix)mForm.Items.Item(MTX_MAIN).Specific;
                oEdtDate = (SAPbouiCOM.Item)mForm.Items.Item(EDT_DATE);
                btnOW = (SAPbouiCOM.Item)mForm.Items.Item(BTN_ORDFAB);
                btnNS = (SAPbouiCOM.Item)mForm.Items.Item(BTN_ENSAMBLAR);
                btnActS = (SAPbouiCOM.Item)mForm.Items.Item(BTN_ACTSERIE); 
                btnAdd = (SAPbouiCOM.Item)mForm.Items.Item(BTN_ADD);
                btnCE = (SAPbouiCOM.Item)mForm.Items.Item(BTN_FILE);
                LoadDefaults();
                //AddRow();
                oMatrix.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                StatusMessageError("cargarOpcionesPorDefecto > " + ex.Message);
            }
            finally { mForm.Freeze(false); }
        }

        private void Automanage()
        {
            mForm.Items.Item("btnPrev").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Visible, -1, BoModeVisualBehavior.mvb_False);
            mForm.Items.Item("btnPrev").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Visible, 1, BoModeVisualBehavior.mvb_True);
        }


        private void LoadDefaults()
        {
            dsHEAD.SetValue("U_EXC_ESTA", 0, "O");
            dsHEAD.SetValue("U_EXC_FEPR", 0, DateTime.Now.ToString(Constants.SAPDATE_FORMAT));
            LoadCombo();
            UpdateCFLConditions();
            btnOW.Enabled = false;
            btnNS.Enabled = false;
            btnActS.Enabled = false;
            btnCE.Enabled = true;
            if (oMatrix.RowCount == 0) dsDETA.Clear();
        }

        private void UpdateCFLConditions()
        {
            try
            {
                SAPbouiCOM.ChooseFromListCollection oChooseFromListCollection = mForm.ChooseFromLists;
                SAPbouiCOM.Conditions oConditions = null;
                SAPbouiCOM.Condition oCondition = null;
                SAPbouiCOM.ChooseFromList oChooseFromList = null;

                //oChooseFromList = oChooseFromListCollection.Item("cflLMAT");
                //oConditions = oChooseFromList.GetConditions();
                ////oCondition = oConditions.Add();
                ////oCondition.Alias = "U_EXX_TIPOEXIS";
                ////oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                ////oCondition.CondVal = "03";
                //oChooseFromList.SetConditions(oConditions);

                oChooseFromList = oChooseFromListCollection.Item("cflCHIP");
                oConditions = oChooseFromList.GetConditions();
                oCondition = oConditions.Add();
                oCondition.Alias = "U_EXC_DISPROD";
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCondition.CondVal = "Y";//SimCards
                oChooseFromList.SetConditions(oConditions);

                //oChooseFromList = oChooseFromListCollection.Item("cflAcc1");
                //oConditions = oChooseFromList.GetConditions();
                //oCondition = oConditions.Add();
                //oCondition.Alias = "U_EXC_DISPROD";
                //oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                //oCondition.CondVal = "Y";//            
                //oChooseFromList.SetConditions(oConditions);

                //oChooseFromList = oChooseFromListCollection.Item("cflAcc2");
                //oConditions = oChooseFromList.GetConditions();
                //oCondition = oConditions.Add();
                //oCondition.Alias = "U_EXC_DISPROD";
                //oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                //oCondition.CondVal = "Y";//            
                //oChooseFromList.SetConditions(oConditions);

            }
            catch (Exception e)
            {
                StatusMessageError("UpdateCFLConditions() > " + e.Message);
            }
        }

        private void LoadCombo()
        {
            try
            {
                //InstanciateCombo((ComboBox)oMatrix.GetCellSpecific(C_RUT_RUTAS, 0), Queries.GetComboRutas());
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public bool HandleItemEvents(SAPbouiCOM.ItemEvent itemEvent)
        {
            var result = true;
            try
            {
                switch (itemEvent.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                        if (!itemEvent.BeforeAction)
                            whenChooseFromList(itemEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                        result = WhenFormLoad(itemEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        result = WhenItemPressed(itemEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                        if (!itemEvent.BeforeAction)
                            result = WhenComboSelectAfter(itemEvent);
                        else
                            result = WhenComboSelectBefore(itemEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_VALIDATE:
                        if (!itemEvent.BeforeAction)
                        {
                            result = WhenValidate(itemEvent);
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                result = false;
                StatusMessageError("HandleItemEvents() > " + ex.Message);
            }
            return result;
        }

        //Método maneja evento
        private bool WhenValidate(SAPbouiCOM.ItemEvent oEvent)
        {
            bool res = true;
            try
            {
                mForm.Freeze(true);
                switch (oEvent.ItemUID)
                {
                    case "mtxMain": MatrixValidate(oEvent); break; //MATRIX

                    default:
                        break;
                }
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                mForm.Freeze(false);
            }

            return res;
        }

        private bool MatrixValidate(ItemEvent oEvent)
        {
            bool res = true;

            try
            {
                string serie = ((EditText)oMatrix.GetCellSpecific("Col_15", oEvent.Row)).Value.ToString();

                mForm.Freeze(true);
                switch (oEvent.ColUID)
                {
                    case "Col_15":

                        if (!string.IsNullOrEmpty(serie) && oEvent.ItemChanged)
                        {
                            oMatrix.FlushToDataSource();

                            res = ValidarDuplicidad(dsDETA.GetValue("U_EXC_CPRO", oEvent.Row - 1), dsDETA.GetValue("U_EXC_NSER", oEvent.Row - 1));
                            if (res)
                            {
                                SetDetalleSerItem(dsDETA.GetValue("U_EXC_CPRO", oEvent.Row - 1), dsDETA.GetValue("U_EXC_NSER", oEvent.Row - 1), oEvent.Row - 1);
                            }
                            else
                            {
                                dsDETA.SetValue("U_EXC_CEQP", oEvent.Row - 1, "");
                                dsDETA.SetValue("U_EXC_DEQP", oEvent.Row - 1, "");
                                dsDETA.SetValue("U_EXC_SEQP", oEvent.Row - 1, "");
                                dsDETA.SetValue("U_EXC_MARC", oEvent.Row - 1, "");
                                dsDETA.SetValue("U_EXC_MODE", oEvent.Row - 1, "");
                                dsDETA.SetValue("U_EXC_CCHI", oEvent.Row - 1, "");
                                dsDETA.SetValue("U_EXC_DCHI", oEvent.Row - 1, "");
                                dsDETA.SetValue("U_EXC_SCHI", oEvent.Row - 1, "");
                                dsDETA.SetValue("U_EXC_IMEI", oEvent.Row - 1, "");
                                dsDETA.SetValue("U_EXC_OPER", oEvent.Row - 1, "");
                                dsDETA.SetValue("U_EXC_ALMI", oEvent.Row - 1, "");
                                dsDETA.SetValue("U_EXC_ALMS", oEvent.Row - 1, "");

                                dsDETA.SetValue("U_EXC_CACC1", oEvent.Row - 1, "");
                                dsDETA.SetValue("U_EXC_DACC1", oEvent.Row - 1, "");
                                dsDETA.SetValue("U_EXC_CACC2", oEvent.Row - 1, "");
                                dsDETA.SetValue("U_EXC_DACC2", oEvent.Row - 1, "");

                                StatusMessageWarning("No puede ponerse serie repetida");
                            }

                            oMatrix.LoadFromDataSource();
                        }
                        break;
                }
                oMatrix.AutoResizeColumns();
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                mForm.Freeze(false);
            }

            return res;
        }

        private bool WhenFormLoad(SAPbouiCOM.ItemEvent oEvent)
        {
            if (!oEvent.BeforeAction) return true;
            try
            {
                mForm.Freeze(true);
            }
            catch (Exception)
            {
                throw;
            }
            finally { mForm.Freeze(false); }

            return true;
        }

        private bool WhenItemPressed(SAPbouiCOM.ItemEvent oEvent)
        {
            bool res = true;
            try
            {
                switch (oEvent.ItemUID)
                {
                    case BTN_OK:
                        if (!oEvent.BeforeAction)
                        {
                            if (mForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && oEvent.ActionSuccess)
                            {

                            }
                        }
                        else
                        {
                            if (dsDETA.Size == 0)
                            {
                                StatusMessageError("Debe tener detalles.");
                                res = false;
                            }
                            res = ValidarReceta(mForm);
                        }
                        if (oEvent.ActionSuccess && mForm.Mode == BoFormMode.fm_OK_MODE)
                        {
                            bool actOW = true;
                            bool actNS = true;
                            bool actAct = true;
                            oMatrix.Item.Enabled = true;
                            oEdtDate.Enabled = true;

                            switch (dsHEAD.GetValue("U_EXC_ESTA", 0))
                            {
                                case "O":
                                    actOW = true;
                                    actNS = false;
                                    actAct = false;
                                    break;
                                case "P":
                                    actOW = false;
                                    actNS = true;
                                    actAct = true;
                                    break;
                                case "F":
                                    oMatrix.Item.Enabled = false;
                                    oEdtDate.Enabled = false;
                                    actOW = false;
                                    actNS = false;
                                    actAct = true;
                                    break;

                                default:
                                    break;
                            }

                            btnAdd.Enabled = false;
                            btnCE.Enabled = false;
                            btnOW.Enabled = actOW;
                            btnNS.Enabled = actNS;
                            btnActS.Enabled = actAct;
                            btnAdd.Enabled = actNS;
                        }
                        else
                        {
                            if (dsDETA.Size == 0)
                            {
                                StatusMessageError("Debe tener detalles.");
                                res = false;
                            }
                        }
                        break;
                    case BTN_ORDFAB:
                        if (!oEvent.BeforeAction)
                        {
                            if (dsDETA.Size == 0)
                            {
                                StatusMessageWarning("Debe tener detalles.");
                                res = false;
                            }
                            if (res) res = ValidarReceta(mForm);
                            if (res) res = ProcesoOrdenFabricacion(oEvent);
                        }
                        break;
                    case BTN_ENSAMBLAR:
                        if (!oEvent.BeforeAction)
                        {
                            if (dsDETA.Size == 0)
                            {
                                StatusMessageWarning("Debe tener detalles.");
                                res = false;
                            }
                            if (res) res = ValidarReceta(mForm);
                            if (res) res = ProcesoDesensamble(oEvent);
                        }
                        break;
                    case BTN_ACTSERIE:
                        if (!oEvent.BeforeAction)
                        {
                            if (dsDETA.Size == 0)
                            {
                                StatusMessageWarning("Debe tener detalles.");
                                res = false;
                            }
                            if (res) res = ActualizarSerie(oEvent);
                        }
                        break;
                    case BTN_ADD:
                        if (!oEvent.BeforeAction)
                            AddRow(MTX_MAIN);
                        break;
                    case BTN_SERIE:
                        if (!oEvent.BeforeAction)
                            Conexion.application.ActivateMenuItem("51229");
                        break;
                    case BTN_FILE:
                        if (oEvent.ActionSuccess)
                            CargarArchivo();
                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            finally { mForm.Freeze(false); }
            return res;
        }

        private bool ValidarReceta(SAPbouiCOM.Form mForm)
        {
            bool valido = true;
            SAPbobsCOM.Recordset oRS = null;

            try
            {
                oRS = (SAPbobsCOM.Recordset)Conexion.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                int count = 0;
                for (int row = 1; row <= oMatrix.RowCount; row++)
                {
                    string ItemCode = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(row).Specific).Value;
                    string Chip = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_7").Cells.Item(row).Specific).Value;
                    string Accesorio1 = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_30").Cells.Item(row).Specific).Value;
                    string Accesorio2 = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_32").Cells.Item(row).Specific).Value;

                    oRS.DoQuery(Queries.GetDetalleProductoProducir(ItemCode));
                    List<string> Receta = new List<string>();
                    while (!oRS.EoF)
                    {
                        Receta.Add(oRS.Fields.Item("Code").Value.ToString());
                        oRS.MoveNext();
                    }

                    count = Receta.Count(x => x.Equals(Chip));
                    if (count == 0) throw new Exception("Línea " + row + ": el chip seleccionado no corresponde a la receta del producto a fabricar.");


                    if (!string.IsNullOrEmpty(Accesorio1))
                    {
                        count = Receta.Count(x => x.Equals(Accesorio1));
                        if (count == 0) throw new Exception("Línea " + row + ": el accesorio 1 seleccionado no corresponde a la receta del producto a fabricar.");
                    }

                    if (!string.IsNullOrEmpty(Accesorio2))
                    {
                        count = Receta.Count(x => x.Equals(Accesorio2));
                        if (count == 0) throw new Exception("Línea " + row + ": el accesorio 2 seleccionado no corresponde a la receta del producto a fabricar.");
                    }
                }
            }
            catch (Exception ex)
            {
                StatusMessageWarning(ex.Message);
                valido = false;
            }
            return valido;
        }

        private bool ProcesoDesensamble(ItemEvent oEvent)
        {
            bool sta = true;
            int res;

            try
            {
                if (mForm.Mode != BoFormMode.fm_ADD_MODE && mForm.Mode != BoFormMode.fm_FIND_MODE)
                {
                    if (dsHEAD.GetValue("U_EXC_ESTA", 0) != "P")
                    {
                        StatusMessageWarning("Primero debe generar la orden de fabricación");
                        return false;
                    }
                    StatusMessageWarning("Iniciando proceso de desensamble.");

                    oMatrix.FlushToDataSource();
                    for (int i = 0; i < dsDETA.Size; i++)
                    {
                        try
                        {
                            if (GetStatusOrder(dsDETA.GetValue("U_EXC_ORDT", i)) != "L")
                            {
                                //string ItemCode = dsDETA.GetValue("U_EXC_CEQP", i);
                                //string SerieNumberEqp = string.IsNullOrEmpty(dsDETA.GetValue("U_EXC_SEQP", i)) ? dsDETA.GetValue("U_EXC_IMEI", i).ToString() : dsDETA.GetValue("U_EXC_SEQP", i);
                                //string Chip = dsDETA.GetValue("U_EXC_CCHI", i);
                                //string SerieNumberCh = GetInternalNumberSerie(dsDETA.GetValue("U_EXC_CCHI", i), dsDETA.GetValue("U_EXC_IMEICH", i));
                                //string Fecha = dsHEAD.GetValue("U_EXC_FEPR", 0).ToString();

                                Conexion.company.StartTransaction();
                                sta = ActualizarOrdProd(i, BoProductionOrderStatusEnum.boposReleased);
                                if (sta) sta = RetirarProducto(i);
                                if (sta) sta = DevolverProducto(i);
                                if (sta) sta = ActualizarOrdProd(i, BoProductionOrderStatusEnum.boposClosed);
                                if (sta && Conexion.company.InTransaction)
                                {
                                    Conexion.company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

                                    #region Actualiza serie por UI
                                    /*
                                    try
                                    {
                                        //Actualiza serie
                                        //Conexion.application.OpenForm(BoFormObjectEnum.fo_SerialNumbersForItems, "", ItemCode); //No funcionó se debe pasar 2 parametros
                                        //SAPbouiCOM.FormCreationParams oFormParams  = Conexion.application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                                        //oFormParams.ObjectType = "10000045";
                                        //oFormParams.FormType = ((int)BoFormObjectEnum.fo_SerialNumbersForItems).ToString();
                                        //oFormParams.UniqueID = dsDETA.GetValue("U_EXC_NSER", i);
                                        //Conexion.application.Forms.AddEx(oFormParams);

                                        Conexion.application.Menus.Item("12034").Activate();
                                        SAPbouiCOM.Form oFormSerie = Conexion.application.Forms.ActiveForm;

                                        oFormSerie.Items.Item("4").Click();
                                        ((EditText)oFormSerie.Items.Item("4").Specific).Value = ItemCode;
                                        ((EditText)oFormSerie.Items.Item("54").Specific).Value = SerieNumberEqp;
                                        oFormSerie.Items.Item("1").Click(BoCellClickType.ct_Regular);
                                        Thread.Sleep(500);
                                        ((EditText)oFormSerie.Items.Item("53").Specific).Value = SerieNumberEqp;
                                        Matrix oMatrix = (Matrix)oFormSerie.Items.Item("43").Specific;
                                        ((EditText)oMatrix.Columns.Item("8").Cells.Item(1).Specific).Value = dsHEAD.GetValue("U_EXC_FEPR", 0);
                                        ((EditText)oMatrix.Columns.Item("7").Cells.Item(1).Specific).Value = dsHEAD.GetValue("U_EXC_FEPR", 0);
                                        //((EditText)oMatrix.Columns.Item("U_EXC_IMEI").Cells.Item(1).Specific).Value = dsDETA.GetValue("U_EXC_IMEI", i);
                                        //((EditText)oMatrix.Columns.Item("U_EXC_MARCA").Cells.Item(1).Specific).Value = dsDETA.GetValue("U_EXC_MARC", i);
                                        //((EditText)oMatrix.Columns.Item("U_EXC_MODELO").Cells.Item(1).Specific).Value = dsDETA.GetValue("U_EXC_MODE", i);
                                        //((EditText)oMatrix.Columns.Item("U_EXC_PRODPOR").Cells.Item(1).Specific).Value = dsDETA.GetValue("U_EXC_PROP", i);
                                        //((EditText)oMatrix.Columns.Item("U_EXC_FIRMW").Cells.Item(1).Specific).Value = dsDETA.GetValue("U_EXC_FWAR", i);
                                        oFormSerie.Items.Item("1").Click(BoCellClickType.ct_Regular);
                                        oFormSerie.Close();
                                        GenericQuery(Queries.UpdateSerieEquipo(ItemCode, SerieNumberEqp, dsDETA, i));

                                        Conexion.application.Menus.Item("12034").Activate();
                                        SAPbouiCOM.Form oFormSerie2 = Conexion.application.Forms.ActiveForm;
                                        oFormSerie2.Items.Item("4").Click();
                                        ((EditText)oFormSerie2.Items.Item("4").Specific).Value = Chip;
                                        ((EditText)oFormSerie2.Items.Item("54").Specific).Value = SerieNumberCh;
                                        oFormSerie2.Items.Item("1").Click(BoCellClickType.ct_Regular);
                                        Thread.Sleep(500);
                                        ((EditText)oFormSerie2.Items.Item("53").Specific).Value = dsDETA.GetValue("U_EXC_IMEICH", i);
                                        oMatrix = (Matrix)oFormSerie2.Items.Item("43").Specific;
                                        ((EditText)oMatrix.Columns.Item("8").Cells.Item(1).Specific).Value = dsHEAD.GetValue("U_EXC_FEPR", 0);
                                        ((EditText)oMatrix.Columns.Item("7").Cells.Item(1).Specific).Value = dsHEAD.GetValue("U_EXC_FEPR", 0);
                                        //((EditText)oMatrix.Columns.Item("U_EXC_SIMCARD").Cells.Item(1).Specific).Value = dsDETA.GetValue("U_EXC_SIMC", i);
                                        //((EditText)oMatrix.Columns.Item("U_EXC_LINTEL").Cells.Item(1).Specific).Value = dsDETA.GetValue("U_EXC_LTEL", i);
                                        //((EditText)oMatrix.Columns.Item("U_EXC_OPERAD").Cells.Item(1).Specific).Value = dsDETA.GetValue("U_EXC_OPER", i);
                                        //((EditText)oMatrix.Columns.Item("U_EXC_PAQDATOS").Cells.Item(1).Specific).Value = dsDETA.GetValue("U_EXC_PQDA", i);
                                        //((EditText)oMatrix.Columns.Item("U_EXC_APN").Cells.Item(1).Specific).Value = dsDETA.GetValue("U_EXC_DAPN", i);
                                        //((EditText)oMatrix.Columns.Item("U_EXC_TIPIP").Cells.Item(1).Specific).Value = dsDETA.GetValue("U_EXC_TIIP", i);
                                        oFormSerie2.Items.Item("1").Click(BoCellClickType.ct_Regular);
                                        oFormSerie2.Close();
                                        GenericQuery(Queries.UpdateSerieChip(Chip, SerieNumberCh, dsDETA, i));
                                    }
                                    catch (Exception ex)
                                    {
                                        //StatusMessageError(string.Format("Falló en actualizar la serie: {0}", ex));
                                        if (ex.Message.Contains("Close"))
                                        {
                                            try
                                            {
                                                Conexion.application.Menus.Item("12034").Activate();
                                                SAPbouiCOM.Form oFormSerie2 = Conexion.application.Forms.ActiveForm;
                                                oFormSerie2.Items.Item("4").Click();
                                                ((EditText)oFormSerie2.Items.Item("4").Specific).Value = Chip;
                                                ((EditText)oFormSerie2.Items.Item("54").Specific).Value = SerieNumberCh;
                                                oFormSerie2.Items.Item("1").Click(BoCellClickType.ct_Regular);

                                                ((EditText)oFormSerie2.Items.Item("53").Specific).Value = dsDETA.GetValue("U_EXC_IMEICH", i); ;
                                                oMatrix = (Matrix)oFormSerie2.Items.Item("43").Specific;
                                                ((EditText)oMatrix.Columns.Item("8").Cells.Item(1).Specific).Value = dsHEAD.GetValue("U_EXC_FEPR", 0);
                                                ((EditText)oMatrix.Columns.Item("7").Cells.Item(1).Specific).Value = dsHEAD.GetValue("U_EXC_FEPR", 0);
                                                ((EditText)oMatrix.Columns.Item("U_EXC_SIMCARD").Cells.Item(1).Specific).Value = dsDETA.GetValue("U_EXC_SIMC", i);
                                                ((EditText)oMatrix.Columns.Item("U_EXC_LINTEL").Cells.Item(1).Specific).Value = dsDETA.GetValue("U_EXC_LTEL", i);
                                                ((EditText)oMatrix.Columns.Item("U_EXC_OPERAD").Cells.Item(1).Specific).Value = dsDETA.GetValue("U_EXC_OPER", i);
                                                ((EditText)oMatrix.Columns.Item("U_EXC_PAQDATOS").Cells.Item(1).Specific).Value = dsDETA.GetValue("U_EXC_PQDA", i);
                                                ((EditText)oMatrix.Columns.Item("U_EXC_APN").Cells.Item(1).Specific).Value = dsDETA.GetValue("U_EXC_DAPN", i);
                                                ((EditText)oMatrix.Columns.Item("U_EXC_TIPIP").Cells.Item(1).Specific).Value = dsDETA.GetValue("U_EXC_TIIP", i);
                                                oFormSerie2.Items.Item("1").Click(BoCellClickType.ct_Regular);
                                                oFormSerie2.Close();
                                            }
                                            catch (Exception)
                                            {
                                            }
                                        }
                                    }
                                    */
                                    #endregion
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            if (Conexion.company.InTransaction) Conexion.company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);

                            if (!ex.Message.Contains("focus"))
                                StatusMessageError(string.Format("Falló en ensanmblar: {0}", ex));
                            sta = false;
                            return sta;
                        }
                        StatusMessageWarning("Generando desensambles " + (i + 1) + " de " + dsDETA.Size);
                    }

                    //   dsHEAD.SetValue("U_EXC_ESTA", 0, "F");
                    //Cambio 20220201
                    oMatrix.LoadFromDataSource();

                    if (sta)
                    {
                        dsHEAD.SetValue("U_EXC_ESTA", 0, "F");
                        if (mForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            mForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                        mForm.Items.Item(BTN_OK).Click();
                        StatusMessageWarning("Generación de desensambles culminó satisfactoriamente.");
                        Conexion.application.ActivateMenuItem(Constants.Actualizar_Registro);
                    }
                }
                else
                {
                    StatusMessageError("El registro debe estar creado para procesar el desensamble");
                    return false;
                }
            }
            catch (Exception ex)
            {
                sta = false;
                StatusMessageError(ex.Message);

                oMatrix.LoadFromDataSource();
                dsHEAD.SetValue("U_EXC_ESTA", 0, "P");
                if (mForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                    mForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                mForm.Items.Item(BTN_OK).Click();
                StatusMessageWarning("Generación de desensambles culminó con errores.");
            }
            return sta;
        }

        private bool ActualizarSerie(ItemEvent oEvent)
        {
            bool sta = true;
            int res;

            try
            {
                if (mForm.Mode != BoFormMode.fm_ADD_MODE && mForm.Mode != BoFormMode.fm_FIND_MODE)
                {
                    if (dsHEAD.GetValue("U_EXC_ESTA", 0) != "P")
                    {
                        StatusMessageWarning("Primero debe generar la orden de fabricación");
                        return false;
                    }
                    StatusMessageWarning("Iniciando actualización de series.");

                    for (int i = 0; i < dsDETA.Size; i++)
                    {
                        string ItemCode = dsDETA.GetValue("U_EXC_CEQP", i);
                        string SerieNumberEqp = string.IsNullOrEmpty(dsDETA.GetValue("U_EXC_SEQP", i)) ? dsDETA.GetValue("U_EXC_IMEI", i).ToString() : dsDETA.GetValue("U_EXC_SEQP", i);

                        try
                        {
                            Conexion.application.Menus.Item("12034").Activate();
                            SAPbouiCOM.Form oFormSerie = Conexion.application.Forms.ActiveForm;

                            oFormSerie.Items.Item("4").Click();
                            ((EditText)oFormSerie.Items.Item("4").Specific).Value = ItemCode;
                            ((EditText)oFormSerie.Items.Item("54").Specific).Value = SerieNumberEqp;
                            oFormSerie.Items.Item("1").Click(BoCellClickType.ct_Regular);

                            ((EditText)oFormSerie.Items.Item("53").Specific).Value = SerieNumberEqp;
                            Matrix oMatrix = (Matrix)oFormSerie.Items.Item("43").Specific;
                            ((EditText)oMatrix.Columns.Item("8").Cells.Item(1).Specific).Value = dsHEAD.GetValue("U_EXC_FEPR", 0);
                            ((EditText)oMatrix.Columns.Item("7").Cells.Item(1).Specific).Value = dsHEAD.GetValue("U_EXC_FEPR", 0);
                            ((EditText)oMatrix.Columns.Item("U_EXC_IMEI").Cells.Item(1).Specific).Value = dsDETA.GetValue("U_EXC_IMEI", i);
                            ((EditText)oMatrix.Columns.Item("U_EXC_MARCA").Cells.Item(1).Specific).Value = dsDETA.GetValue("U_EXC_MARC", i);
                            ((EditText)oMatrix.Columns.Item("U_EXC_MODELO").Cells.Item(1).Specific).Value = dsDETA.GetValue("U_EXC_MODE", i);
                            ((EditText)oMatrix.Columns.Item("U_EXC_PRODPOR").Cells.Item(1).Specific).Value = dsDETA.GetValue("U_EXC_PROP", i);
                            ((EditText)oMatrix.Columns.Item("U_EXC_FIRMW").Cells.Item(1).Specific).Value = dsDETA.GetValue("U_EXC_FWAR", i);
                            oFormSerie.Items.Item("1").Click(BoCellClickType.ct_Regular);
                            oFormSerie.Close();
                        }
                        catch (Exception ex)
                        {
                        }
                    }

                    for (int i = 0; i < dsDETA.Size; i++)
                    {
                        string Chip = dsDETA.GetValue("U_EXC_CCHI", i);
                        string SerieNumberCh = GetInternalNumberSerie(dsDETA.GetValue("U_EXC_CCHI", i), dsDETA.GetValue("U_EXC_IMEICH", i));

                        try
                        {
                            Conexion.application.Menus.Item("12034").Activate();
                            SAPbouiCOM.Form oFormSerie2 = Conexion.application.Forms.ActiveForm;
                            oFormSerie2.Items.Item("4").Click();
                            ((EditText)oFormSerie2.Items.Item("4").Specific).Value = Chip;
                            ((EditText)oFormSerie2.Items.Item("54").Specific).Value = SerieNumberCh;
                            oFormSerie2.Items.Item("1").Click(BoCellClickType.ct_Regular);

                            ((EditText)oFormSerie2.Items.Item("53").Specific).Value = dsDETA.GetValue("U_EXC_IMEICH", i);
                            oMatrix = (Matrix)oFormSerie2.Items.Item("43").Specific;
                            ((EditText)oMatrix.Columns.Item("8").Cells.Item(1).Specific).Value = dsHEAD.GetValue("U_EXC_FEPR", 0);
                            ((EditText)oMatrix.Columns.Item("7").Cells.Item(1).Specific).Value = dsHEAD.GetValue("U_EXC_FEPR", 0);
                            ((EditText)oMatrix.Columns.Item("U_EXC_SIMCARD").Cells.Item(1).Specific).Value = dsDETA.GetValue("U_EXC_SIMC", i);
                            ((EditText)oMatrix.Columns.Item("U_EXC_LINTEL").Cells.Item(1).Specific).Value = dsDETA.GetValue("U_EXC_LTEL", i);
                            ((EditText)oMatrix.Columns.Item("U_EXC_OPERAD").Cells.Item(1).Specific).Value = dsDETA.GetValue("U_EXC_OPER", i);
                            ((EditText)oMatrix.Columns.Item("U_EXC_PAQDATOS").Cells.Item(1).Specific).Value = dsDETA.GetValue("U_EXC_PQDA", i);
                            ((EditText)oMatrix.Columns.Item("U_EXC_APN").Cells.Item(1).Specific).Value = dsDETA.GetValue("U_EXC_DAPN", i);
                            ((EditText)oMatrix.Columns.Item("U_EXC_TIPIP").Cells.Item(1).Specific).Value = dsDETA.GetValue("U_EXC_TIIP", i);
                            oFormSerie2.Items.Item("1").Click(BoCellClickType.ct_Regular);
                            oFormSerie2.Close();
                        }
                        catch (Exception ex)
                        {
                        }
                    }

                    StatusMessageWarning("Generación de desensambles culminó satisfactoriamente.");
                }
            }
            catch (Exception ex)
            {
                StatusMessageError(ex.Message);
            }
            return true;
        }

        private bool RetirarProducto(int row)
        {
            Documents oRetiroProducto = null;
            bool sta = true;
            int res;
            try
            {
                oRetiroProducto = (Documents)Conexion.company.GetBusinessObject(BoObjectTypes.oInventoryGenExit);
                oRetiroProducto.DocDate = DateTime.ParseExact((dsHEAD.GetValue("U_EXC_FEPR", 0)), "yyyyMMdd", CultureInfo.InvariantCulture);
                oRetiroProducto.UserFields.Fields.Item("U_EXX_TIPOOPER").Value = AddonProduccionEnsDes.Properties.Resources.TOperacionSalida;

                oRetiroProducto.Lines.BaseType = (int)BoObjectTypes.oProductionOrders;
                oRetiroProducto.Lines.BaseEntry = int.Parse(dsDETA.GetValue("U_EXC_ORDT", row));
                oRetiroProducto.Lines.Quantity = 1;
                oRetiroProducto.Lines.SerialNumbers.SystemSerialNumber = int.Parse(dsDETA.GetValue("U_EXC_NSER", row));
                oRetiroProducto.Lines.SerialNumbers.Quantity = 1;
                res = oRetiroProducto.Add();
                if (res != 0)
                {
                    if (Conexion.company.InTransaction) Conexion.company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    StatusMessageError(string.Format("Falló la fabricación {0}:{1}", "", Conexion.company.GetLastErrorDescription()));
                    sta = false;
                }
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                LiberarObjetoGenerico(oRetiroProducto);
            }
            return sta;
        }

        private bool DevolverProducto(int row)
        {
            Documents oReciboProducto = null;
            bool sta = true;
            int res;
            try
            {
                oReciboProducto = (Documents)Conexion.company.GetBusinessObject(BoObjectTypes.oInventoryGenEntry);
                oReciboProducto.DocDate = DateTime.ParseExact((dsHEAD.GetValue("U_EXC_FEPR", 0)), "yyyyMMdd", CultureInfo.InvariantCulture);
                oReciboProducto.UserFields.Fields.Item("U_EXX_TIPOOPER").Value = AddonProduccionEnsDes.Properties.Resources.TOperacionEntrada;

                int index = 0;
                string Serie = string.IsNullOrEmpty(dsDETA.GetValue("U_EXC_SEQP", row)) ? dsDETA.GetValue("U_EXC_IMEI", row).ToString() : dsDETA.GetValue("U_EXC_SEQP", row);

                oReciboProducto.Lines.SetCurrentLine(index);
                oReciboProducto.Lines.BaseType = (int)BoObjectTypes.oProductionOrders;
                oReciboProducto.Lines.BaseEntry = int.Parse(dsDETA.GetValue("U_EXC_ORDT", row));
                oReciboProducto.Lines.BaseLine = index;
                oReciboProducto.Lines.Quantity = 1;
                oReciboProducto.Lines.WarehouseCode = dsDETA.GetValue("U_EXC_ALMI", row);
                oReciboProducto.Lines.SerialNumbers.InternalSerialNumber = Serie; // (dsDETA.GetValue("U_EXC_SEQP", row));
                oReciboProducto.Lines.SerialNumbers.ManufacturerSerialNumber = Serie; // (dsDETA.GetValue("U_EXC_SEQP", row));
                oReciboProducto.Lines.SerialNumbers.ReceptionDate = DateTime.ParseExact((dsHEAD.GetValue("U_EXC_FEPR", 0)), "yyyyMMdd", CultureInfo.InvariantCulture);
                oReciboProducto.Lines.SerialNumbers.ManufactureDate = DateTime.ParseExact((dsHEAD.GetValue("U_EXC_FEPR", 0)), "yyyyMMdd", CultureInfo.InvariantCulture);

                //Nuevo 202202003
                oReciboProducto.Lines.SerialNumbers.UserFields.Fields.Item("U_EXC_MARCA").Value = dsDETA.GetValue("U_EXC_MARC", row);
                oReciboProducto.Lines.SerialNumbers.UserFields.Fields.Item("U_EXC_MODELO").Value = dsDETA.GetValue("U_EXC_MODE", row);
                oReciboProducto.Lines.SerialNumbers.UserFields.Fields.Item("U_EXC_IMEI").Value = dsDETA.GetValue("U_EXC_IMEI", row);
                oReciboProducto.Lines.SerialNumbers.UserFields.Fields.Item("U_EXC_FIRMW").Value = dsDETA.GetValue("U_EXC_FWAR", row);
                oReciboProducto.Lines.SerialNumbers.UserFields.Fields.Item("U_EXC_FOTA").Value = dsDETA.GetValue("U_EXC_FOTA", row);
                oReciboProducto.Lines.SerialNumbers.UserFields.Fields.Item("U_EXC_PRODPOR").Value = dsDETA.GetValue("U_EXC_PROP", row);
                oReciboProducto.Lines.SerialNumbers.Quantity = 1;
                index++;

                oReciboProducto.Lines.Add();
                oReciboProducto.Lines.SetCurrentLine(index);
                oReciboProducto.Lines.BaseType = (int)BoObjectTypes.oProductionOrders;
                oReciboProducto.Lines.BaseEntry = int.Parse(dsDETA.GetValue("U_EXC_ORDT", row));
                oReciboProducto.Lines.BaseLine = index;
                oReciboProducto.Lines.Quantity = 1;
                oReciboProducto.Lines.WarehouseCode = dsDETA.GetValue("U_EXC_ALMI", row);  //AddonProduccionEnsDes.Properties.Resources.AlmSalida;
                oReciboProducto.Lines.SerialNumbers.InternalSerialNumber = (dsDETA.GetValue("U_EXC_LTEL", row));
                oReciboProducto.Lines.SerialNumbers.ManufacturerSerialNumber = (dsDETA.GetValue("U_EXC_SIMC", row));
                oReciboProducto.Lines.SerialNumbers.ReceptionDate = DateTime.ParseExact((dsHEAD.GetValue("U_EXC_FEPR", 0)), "yyyyMMdd", CultureInfo.InvariantCulture);
                oReciboProducto.Lines.SerialNumbers.ManufactureDate = DateTime.ParseExact((dsHEAD.GetValue("U_EXC_FEPR", 0)), "yyyyMMdd", CultureInfo.InvariantCulture);
                oReciboProducto.Lines.SerialNumbers.UserFields.Fields.Item("U_EXC_APN").Value = dsDETA.GetValue("U_EXC_DAPN", row);
                oReciboProducto.Lines.SerialNumbers.UserFields.Fields.Item("U_EXC_TIPIP").Value = dsDETA.GetValue("U_EXC_TIIP", row);
                oReciboProducto.Lines.SerialNumbers.UserFields.Fields.Item("U_EXC_IP").Value = dsDETA.GetValue("U_EXC_NRIP", row);
                oReciboProducto.Lines.SerialNumbers.UserFields.Fields.Item("U_EXC_IMEI").Value = dsDETA.GetValue("U_EXC_SIMC", row);
                oReciboProducto.Lines.SerialNumbers.UserFields.Fields.Item("U_EXC_SIMCARD").Value = dsDETA.GetValue("U_EXC_SIMC", row);
                oReciboProducto.Lines.SerialNumbers.UserFields.Fields.Item("U_EXC_LINEA").Value = dsDETA.GetValue("U_EXC_LINE", row);
                oReciboProducto.Lines.SerialNumbers.UserFields.Fields.Item("U_EXC_PAQDATOS").Value = dsDETA.GetValue("U_EXC_PQDA", row);
                oReciboProducto.Lines.SerialNumbers.UserFields.Fields.Item("U_EXC_PRODPOR").Value = dsDETA.GetValue("U_EXC_PROP", row);
                //oReciboProducto.Lines.SerialNumbers.UserFields.Fields.Item("U_EXC_IMEI").Value = dsDETA.GetValue("U_EXC_IMEICH", row);
                index++;

                //ACCESORIOS
                if (!string.IsNullOrEmpty(dsDETA.GetValue("U_EXC_CACC1", row)))
                {
                    oReciboProducto.Lines.Add();
                    oReciboProducto.Lines.SetCurrentLine(index);
                    oReciboProducto.Lines.BaseType = (int)BoObjectTypes.oProductionOrders;
                    oReciboProducto.Lines.BaseEntry = int.Parse(dsDETA.GetValue("U_EXC_ORDT", row));
                    oReciboProducto.Lines.BaseLine = index;
                    oReciboProducto.Lines.Quantity = 1;
                    oReciboProducto.Lines.WarehouseCode = dsDETA.GetValue("U_EXC_ALMI", row); //AddonProduccionEnsDes.Properties.Resources.AlmSalida;
                    index++;
                }
                if (!string.IsNullOrEmpty(dsDETA.GetValue("U_EXC_CACC2", row)))
                {
                    oReciboProducto.Lines.Add();
                    oReciboProducto.Lines.SetCurrentLine(index);
                    oReciboProducto.Lines.BaseType = (int)BoObjectTypes.oProductionOrders;
                    oReciboProducto.Lines.BaseEntry = int.Parse(dsDETA.GetValue("U_EXC_ORDT", row));
                    oReciboProducto.Lines.BaseLine = index;
                    oReciboProducto.Lines.Quantity = 1;
                    oReciboProducto.Lines.WarehouseCode = dsDETA.GetValue("U_EXC_ALMI", row); //AddonProduccionEnsDes.Properties.Resources.AlmSalida;
                }

                oReciboProducto.Lines.SerialNumbers.Quantity = 1;
                res = oReciboProducto.Add();

                if (res != 0)
                {
                    if (Conexion.company.InTransaction) Conexion.company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    StatusMessageError(string.Format("Falló la fabricación {0}:{1}", "", Conexion.company.GetLastErrorDescription()));
                    sta = false;
                }

                int EquipoSerie = GetAbsEntrySerie(dsDETA.GetValue("U_EXC_CEQP", row), dsDETA.GetValue("U_EXC_IMEI", row).ToString());
                if (EquipoSerie > 0)
                {
                    CompanyService oService = Conexion.company.GetCompanyService();
                    SerialNumberDetailsService oSerialNumbersService = (SerialNumberDetailsService)oService.GetBusinessService(ServiceTypes.SerialNumberDetailsService);
                    SerialNumberDetailParams oSerialNumberDetailParams = (SerialNumberDetailParams)oSerialNumbersService.GetDataInterface(SerialNumberDetailsServiceDataInterfaces.sndsSerialNumberDetailParams);
                    oSerialNumberDetailParams.DocEntry = EquipoSerie;

                    SerialNumberDetail oSerialNumberDetail = oSerialNumbersService.Get(oSerialNumberDetailParams);
                    oSerialNumberDetail.MfrSerialNo = Serie; 
                    oSerialNumberDetail.AdmissionDate = DateTime.ParseExact((dsHEAD.GetValue("U_EXC_FEPR", 0)), "yyyyMMdd", CultureInfo.InvariantCulture);
                    oSerialNumberDetail.ManufacturingDate = DateTime.ParseExact((dsHEAD.GetValue("U_EXC_FEPR", 0)), "yyyyMMdd", CultureInfo.InvariantCulture);
                    oSerialNumberDetail.UserFields.Item("U_EXC_MARCA").Value = dsDETA.GetValue("U_EXC_MARC", row);
                    oSerialNumberDetail.UserFields.Item("U_EXC_MODELO").Value = dsDETA.GetValue("U_EXC_MODE", row);
                    oSerialNumberDetail.UserFields.Item("U_EXC_IMEI").Value = dsDETA.GetValue("U_EXC_IMEI", row);
                    oSerialNumberDetail.UserFields.Item("U_EXC_FIRMW").Value = dsDETA.GetValue("U_EXC_FWAR", row);
                    oSerialNumberDetail.UserFields.Item("U_EXC_FOTA").Value = dsDETA.GetValue("U_EXC_FOTA", row);
                    oSerialNumberDetail.UserFields.Item("U_EXC_PRODPOR").Value = dsDETA.GetValue("U_EXC_PROP", row);
                    oSerialNumbersService.Update(oSerialNumberDetail);
                }

                int ChipSerie = GetAbsEntrySerie(dsDETA.GetValue("U_EXC_CCHI", row), dsDETA.GetValue("U_EXC_IMEICH", row));
                if (ChipSerie > 0)
                {
                    CompanyService oService = Conexion.company.GetCompanyService();
                    SerialNumberDetailsService oSerialNumbersService = (SerialNumberDetailsService)oService.GetBusinessService(ServiceTypes.SerialNumberDetailsService);
                    SerialNumberDetailParams oSerialNumberDetailParams = (SerialNumberDetailParams)oSerialNumbersService.GetDataInterface(SerialNumberDetailsServiceDataInterfaces.sndsSerialNumberDetailParams);
                    oSerialNumberDetailParams.DocEntry = ChipSerie;

                    SerialNumberDetail oSerialNumberDetail = oSerialNumbersService.Get(oSerialNumberDetailParams);
                    oSerialNumberDetail.MfrSerialNo = (dsDETA.GetValue("U_EXC_SIMC", row));
                    oSerialNumberDetail.AdmissionDate = DateTime.ParseExact((dsHEAD.GetValue("U_EXC_FEPR", 0)), "yyyyMMdd", CultureInfo.InvariantCulture);
                    oSerialNumberDetail.ManufacturingDate = DateTime.ParseExact((dsHEAD.GetValue("U_EXC_FEPR", 0)), "yyyyMMdd", CultureInfo.InvariantCulture);
                    oSerialNumberDetail.UserFields.Item("U_EXC_APN").Value = dsDETA.GetValue("U_EXC_DAPN", row);
                    oSerialNumberDetail.UserFields.Item("U_EXC_TIPIP").Value = dsDETA.GetValue("U_EXC_TIIP", row);
                    oSerialNumberDetail.UserFields.Item("U_EXC_IP").Value = dsDETA.GetValue("U_EXC_NRIP", row);
                    oSerialNumberDetail.UserFields.Item("U_EXC_IMEI").Value = dsDETA.GetValue("U_EXC_SIMC", row);
                    oSerialNumberDetail.UserFields.Item("U_EXC_SIMCARD").Value = dsDETA.GetValue("U_EXC_SIMC", row);
                    oSerialNumberDetail.UserFields.Item("U_EXC_LINEA").Value = dsDETA.GetValue("U_EXC_LINE", row);
                    oSerialNumberDetail.UserFields.Item("U_EXC_PAQDATOS").Value = dsDETA.GetValue("U_EXC_PQDA", row);
                    oSerialNumberDetail.UserFields.Item("U_EXC_PRODPOR").Value = dsDETA.GetValue("U_EXC_PROP", row);
                    oSerialNumberDetail.UserFields.Item("U_EXC_LINTEL").Value = dsDETA.GetValue("U_EXC_LTEL", row);
                    oSerialNumberDetail.UserFields.Item("U_EXC_OPERAD").Value = dsDETA.GetValue("U_EXC_OPER", row);
                    oSerialNumbersService.Update(oSerialNumberDetail);
                }
            }
            catch (Exception ex)
            {
                sta = false;
                throw ex;
            }
            finally
            {
                LiberarObjetoGenerico(oReciboProducto);
            }
            return sta;
        }

        private double GetPriceFromItem(string itemCode)
        {
            double precio = 0.0;
            SAPbobsCOM.Items oArticulo = null;

            try
            {
                oArticulo = (SAPbobsCOM.Items)Conexion.company.GetBusinessObject(BoObjectTypes.oItems);
                if (!oArticulo.GetByKey(itemCode))
                    throw new Exception($"Artículo {itemCode} no se encuentra registrado en la sociedad");

                precio = Convert.ToDouble(oArticulo.UserFields.Fields.Item("U_EXC_COSTDES").Value);
            }
            catch (Exception)
            {
                throw;
            }
            finally
            { LiberarObjetoGenerico(oArticulo); }

            return precio;
        }

        private string GetStatusOrder(string docEntry)
        {
            SAPbobsCOM.Recordset oRS = null;
            string estado = string.Empty;
            try
            {
                oRS = (SAPbobsCOM.Recordset)Conexion.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRS.DoQuery(Queries.GetStatusOrder(docEntry));

                if (oRS.RecordCount > 0)
                    estado = oRS.Fields.Item("Status").Value;
            }
            catch (Exception)
            {
                throw;
            }
            finally
            { LiberarObjetoGenerico(oRS); }

            return estado;
        }

        private bool ActualizarOrdProd(int row, BoProductionOrderStatusEnum status)
        {
            ProductionOrders oOrdeProd = null;
            bool sta = true;
            int res;
            try
            {
                oOrdeProd = (ProductionOrders)Conexion.company.GetBusinessObject(BoObjectTypes.oProductionOrders);
                oOrdeProd.GetByKey(int.Parse(dsDETA.GetValue("U_EXC_ORDT", row)));
                oOrdeProd.ProductionOrderStatus = status;
                res = oOrdeProd.Update();
                if (res != 0)
                {
                    if (Conexion.company.InTransaction) Conexion.company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    StatusMessageError(string.Format("Falló la fabricación {0}:{1}", "", Conexion.company.GetLastErrorDescription()));
                    sta = false;
                }
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                LiberarObjetoGenerico(oOrdeProd);
            }
            return sta;
        }

        private bool ProcesoOrdenFabricacion(ItemEvent oEvent)
        {
            bool sta = true;
            int res;

            try
            {
                if (mForm.Mode != BoFormMode.fm_ADD_MODE && mForm.Mode != BoFormMode.fm_FIND_MODE)
                {
                    ProductionOrders oProduction = null;
                    StatusMessageWarning("Iniciando proceso de generación orden de fabricaciòn.");
                    oMatrix.FlushToDataSource();

                    for (int i = 0; i < dsDETA.Size; i++)
                    {
                        if (string.IsNullOrEmpty(dsDETA.GetValue("U_EXC_ORDT", i)))
                        {
                            oProduction = (ProductionOrders)Conexion.company.GetBusinessObject(BoObjectTypes.oProductionOrders);
                            oProduction.ProductionOrderType = BoProductionOrderTypeEnum.bopotDisassembly;
                            oProduction.DueDate = DateTime.ParseExact((dsHEAD.GetValue("U_EXC_FEPR", 0)), "yyyyMMdd", CultureInfo.InvariantCulture);
                            oProduction.ItemNo = dsDETA.GetValue("U_EXC_CPRO", i);
                            oProduction.PlannedQuantity = 1;
                            oProduction.Warehouse = dsDETA.GetValue("U_EXC_ALMS", i);

                            oProduction.Lines.ItemNo = dsDETA.GetValue("U_EXC_CEQP", i);
                            oProduction.Lines.Warehouse = dsDETA.GetValue("U_EXC_ALMI", i);
                            oProduction.Lines.ProductionOrderIssueType = BoIssueMethod.im_Manual;
                            oProduction.Lines.Add();

                            oProduction.Lines.ItemNo = dsDETA.GetValue("U_EXC_CCHI", i);
                            oProduction.Lines.Warehouse = AddonProduccionEnsDes.Properties.Resources.AlmSalida;// dsDETA.GetValue("U_EXC_ALMI", i);
                            oProduction.Lines.ProductionOrderIssueType = BoIssueMethod.im_Manual;
                            oProduction.Lines.Add();

                            //ACCESORIOS
                            if (!string.IsNullOrEmpty(dsDETA.GetValue("U_EXC_CACC1", i)))
                            {
                                oProduction.Lines.ItemNo = dsDETA.GetValue("U_EXC_CACC1", i);
                                oProduction.Lines.Warehouse = AddonProduccionEnsDes.Properties.Resources.AlmSalida; //dsDETA.GetValue("U_EXC_ALMI", 0);
                                oProduction.Lines.ProductionOrderIssueType = BoIssueMethod.im_Manual;
                                oProduction.Lines.Add();
                            }
                            if (!string.IsNullOrEmpty(dsDETA.GetValue("U_EXC_CACC2", i)))
                            {
                                oProduction.Lines.ItemNo = dsDETA.GetValue("U_EXC_CACC2", i);
                                oProduction.Lines.Warehouse = AddonProduccionEnsDes.Properties.Resources.AlmSalida; //dsDETA.GetValue("U_EXC_ALMI", 0);
                                oProduction.Lines.ProductionOrderIssueType = BoIssueMethod.im_Manual;
                                oProduction.Lines.Add();
                            }
                            res = oProduction.Add();

                            if (res != 0)
                            {
                                StatusMessageError(string.Format("Falló la fabricación {0}:{1}", "", Conexion.company.GetLastErrorDescription()));
                                sta = false;
                            }
                            else
                            {
                                string DocEntry = Conexion.company.GetNewObjectKey();
                                dsDETA.SetValue("U_EXC_ORDT", i, Conexion.company.GetNewObjectKey());
                                ((EditText)oMatrix.Columns.Item("Col_16").Cells.Item(i + 1).Specific).Value = DocEntry;
                            }
                        }
                        StatusMessageWarning("Creando órdenes de fabricación " + (i + 1) + " de " + dsDETA.Size);
                    }
                    oMatrix.LoadFromDataSource();

                    //Nuevo 20220201
                    if (sta)
                        dsHEAD.SetValue("U_EXC_ESTA", 0, "P");

                    if (mForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        mForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    mForm.Items.Item(BTN_OK).Click();
                    StatusMessageWarning("Generación de órdenes de fabricación culminó satisfactoriamente.");
                }
                else
                {
                    StatusMessageError("El registro debe estar creado para procesar las órdenes de fabricación");
                    return false;
                }
            }
            catch (Exception ex)
            {
                sta = false;
                StatusMessageError(ex.Message);

                oMatrix.LoadFromDataSource();
                if (mForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                    mForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                mForm.Items.Item(BTN_OK).Click();
                StatusMessageWarning("Generación de órdenes de fabricación culminó con errores.");
            }
            finally
            {
                mForm.Freeze(false);
            }

            return sta;
        }

        private bool WhenComboSelectBefore(SAPbouiCOM.ItemEvent oEvent)
        {
            bool res = true;
            try
            {
                switch (oEvent.ItemUID)
                {
                    default:

                        break;
                }
            }
            catch (Exception)
            {
                throw;
            }
            return res;
        }
        private bool WhenComboSelectAfter(SAPbouiCOM.ItemEvent oEvent)
        {
            bool res = true;
            try
            {
                mForm.Freeze(true);
                switch (oEvent.ItemUID)
                {
                    case MTX_MAIN:
                        res = WhenComboMatrixSelectAfter(oEvent);
                        break;
                }
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                mForm.Freeze(false);
            }
            return res;
        }

        private bool WhenComboMatrixSelectAfter(SAPbouiCOM.ItemEvent oEvent)
        {
            bool res = true;
            try
            {
                mForm.Freeze(true);
                switch (oEvent.ColUID)
                {
                    case "Col_15":
                        oMatrix.FlushToDataSource();

                        res = ValidarDuplicidad(dsDETA.GetValue("U_EXC_CPRO", oEvent.Row - 1), dsDETA.GetValue("U_EXC_NSER", oEvent.Row - 1));
                        if (res)
                        {
                            SetDetalleSerItem(dsDETA.GetValue("U_EXC_CPRO", oEvent.Row - 1), dsDETA.GetValue("U_EXC_NSER", oEvent.Row - 1), oEvent.Row - 1);
                        }
                        else
                        {
                            dsDETA.SetValue("U_EXC_CEQP", oEvent.Row - 1, "");
                            dsDETA.SetValue("U_EXC_DEQP", oEvent.Row - 1, "");
                            dsDETA.SetValue("U_EXC_SEQP", oEvent.Row - 1, "");
                            dsDETA.SetValue("U_EXC_MARC", oEvent.Row - 1, "");
                            dsDETA.SetValue("U_EXC_MODE", oEvent.Row - 1, "");
                            dsDETA.SetValue("U_EXC_CCHI", oEvent.Row - 1, "");
                            dsDETA.SetValue("U_EXC_DCHI", oEvent.Row - 1, "");
                            dsDETA.SetValue("U_EXC_SCHI", oEvent.Row - 1, "");
                            dsDETA.SetValue("U_EXC_IMEI", oEvent.Row - 1, "");
                            dsDETA.SetValue("U_EXC_OPER", oEvent.Row - 1, "");
                            dsDETA.SetValue("U_EXC_ALMI", oEvent.Row - 1, "");
                            dsDETA.SetValue("U_EXC_ALMS", oEvent.Row - 1, "");
                            StatusMessageWarning("No puede ponerse serie repetida");
                        }

                        oMatrix.LoadFromDataSource();

                        break;

                }

                oMatrix.AutoResizeColumns();
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                mForm.Freeze(false);
            }
            return res;
        }

        private bool ValidarDuplicidad(string ItemCode, string seriesId)
        {
            bool res = true;
            int count = int.Parse(GenericQuery(Queries.CheckPreviousSeriesProd(ItemCode, seriesId)));
            for (int i = 0; i < dsDETA.Size; i++)
            {
                if (dsDETA.GetValue("U_EXC_NSER", i) == seriesId && dsDETA.GetValue("U_EXC_CPRO", i) == ItemCode) count++;
                if (count > 1) return false;
            }

            return res;
        }

        private bool WhenDataAdd(SAPbouiCOM.ItemEvent oEvent)
        {
            bool res = true;
            switch (oEvent.ItemUID)
            {
            }
            return res;
        }
        private bool WhenDataLoad(SAPbouiCOM.BusinessObjectInfo oEvent)
        {
            bool res = true;
            bool actOW = true;
            bool actNS = true;
            bool actAct = true;
            oMatrix.Item.Enabled = true;
            oEdtDate.Enabled = true;

            switch (dsHEAD.GetValue("U_EXC_ESTA", 0))
            {
                case "O":
                    actOW = true;
                    actNS = false;
                    actAct = false;
                    break;
                case "P":
                    actOW = false;
                    actNS = true;
                    actAct = true;
                    break;
                case "F":
                    oMatrix.Item.Enabled = false;
                    oEdtDate.Enabled = false;
                    actOW = false;
                    actNS = false;
                    actAct = true;
                    break;

                default:
                    break;
            }

            btnAdd.Enabled = false;
            btnCE.Enabled = false;
            btnOW.Enabled = actOW;
            btnNS.Enabled = actNS;
            btnAdd.Enabled = actNS;
            btnActS.Enabled = actAct;
            return res;
        }
        #endregion
        private void whenChooseFromList(SAPbouiCOM.ItemEvent oEvent)
        {
            try
            {
                SAPbouiCOM.IChooseFromListEvent oChooseFromListEvent = (SAPbouiCOM.IChooseFromListEvent)oEvent;
                SAPbouiCOM.DataTable oDataTable = oChooseFromListEvent.SelectedObjects;

                if (oDataTable != null)
                {
                    switch (oEvent.ItemUID)
                    {
                        case MTX_MAIN:
                            whenMatrixChooseFromList(oEvent);
                            break;

                        default:
                            break;
                    }
                }
            }
            catch (Exception e)
            {
                StatusMessageError("whenChooseFromList > " + e.Message);
            }
        }
        private void whenMatrixChooseFromList(SAPbouiCOM.ItemEvent oEvent)
        {
            try
            {
                SAPbouiCOM.IChooseFromListEvent oChooseFromListEvent = (SAPbouiCOM.IChooseFromListEvent)oEvent;
                SAPbouiCOM.DataTable oDataTable = oChooseFromListEvent.SelectedObjects;

                if (oDataTable != null)
                {
                    switch (oEvent.ColUID)
                    {
                        case "Col_0":
                            oMatrix.FlushToDataSource();
                            dsDETA.SetValue("U_EXC_CPRO", oEvent.Row - 1, oDataTable.GetValue("Code", 0).ToString().Trim());
                            dsDETA.SetValue("U_EXC_DPRO", oEvent.Row - 1, oDataTable.GetValue("Name", 0).ToString().Trim());
                            //InstanciateCombo((ComboBox)oMatrix.GetCellSpecific("Col_15", 0), Queries.GetSerieItem(oDataTable.GetValue("Code", 0).Trim()));
                            SetDetalleProducto(oDataTable.GetValue("Code", 0).ToString().Trim(), oEvent.Row - 1);
                            oMatrix.LoadFromDataSource();
                            break;
                        case "Col_7":
                            oMatrix.FlushToDataSource();
                            dsDETA.SetValue("U_EXC_CCHI", oEvent.Row - 1, oDataTable.GetValue("ItemCode", 0).ToString().Trim());
                            dsDETA.SetValue("U_EXC_DCHI", oEvent.Row - 1, oDataTable.GetValue("ItemName", 0).ToString().Trim());
                            //dsDETA.SetValue("U_EXC_SCHI", oEvent.Row - 1, Queries.GetSerieItemDS(oDataTable.GetValue("ItemCode", 0).Trim()));
                            //InstanciateCombo((ComboBox)oMatrix.GetCellSpecific("Col_9", 0), Queries.GetSerieItem(oDataTable.GetValue("ItemCode", 0).Trim()));
                            oMatrix.LoadFromDataSource();
                            break;
                        case "Col_12":
                            oMatrix.FlushToDataSource();
                            dsDETA.SetValue("U_EXC_ALMI", oEvent.Row - 1, oDataTable.GetValue("WhsCode", 0).ToString().Trim());
                            oMatrix.LoadFromDataSource();
                            break;
                        case "Col_13":
                            oMatrix.FlushToDataSource();
                            dsDETA.SetValue("U_EXC_ALMS", oEvent.Row - 1, oDataTable.GetValue("WhsCode", 0).ToString().Trim());
                            oMatrix.LoadFromDataSource();
                            break;
                        case "Col_30": //ACCESORIO
                            oMatrix.FlushToDataSource();
                            if (string.IsNullOrEmpty(oDataTable.GetValue("ItemCode", 0).ToString().Trim()))
                            {
                                dsDETA.SetValue("U_EXC_CACC1", oEvent.Row - 1, "");
                                dsDETA.SetValue("U_EXC_DACC1", oEvent.Row - 1, "");
                            }
                            else
                            {
                                dsDETA.SetValue("U_EXC_CACC1", oEvent.Row - 1, oDataTable.GetValue("ItemCode", 0).ToString().Trim());
                                dsDETA.SetValue("U_EXC_DACC1", oEvent.Row - 1, oDataTable.GetValue("ItemName", 0).ToString().Trim());
                            }
                            oMatrix.LoadFromDataSource();
                            break;
                        case "Col_32": //ACCESORIO 2
                            oMatrix.FlushToDataSource();
                            if (string.IsNullOrEmpty(oDataTable.GetValue("ItemCode", 0).ToString().Trim()))
                            {
                                dsDETA.SetValue("U_EXC_CACC2", oEvent.Row - 1, "");
                                dsDETA.SetValue("U_EXC_DACC2", oEvent.Row - 1, "");
                            }
                            else
                            {
                                dsDETA.SetValue("U_EXC_CACC2", oEvent.Row - 1, oDataTable.GetValue("ItemCode", 0).ToString().Trim());
                                dsDETA.SetValue("U_EXC_DACC2", oEvent.Row - 1, oDataTable.GetValue("ItemName", 0).ToString().Trim());
                            }
                            oMatrix.LoadFromDataSource();
                            break;
                        default:
                            break;
                    }
                }

                oMatrix.AutoResizeColumns();
            }
            catch (Exception e)
            {
                StatusMessageError("whenChooseFromList > " + e.Message);
            }
        }

        private void SetDetalleProducto(string ItemCode, int row)
        {
            SAPbobsCOM.Recordset oRS = null;
            try
            {
                oRS = (SAPbobsCOM.Recordset)Conexion.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRS.DoQuery(Queries.GetDetalleProductoEnsamble(ItemCode));

                if (oRS.RecordCount > 0)
                {
                    dsDETA.SetValue("U_EXC_CEQP", row, oRS.Fields.Item("Code").Value.ToString());
                    dsDETA.SetValue("U_EXC_DEQP", row, oRS.Fields.Item("ItemName").Value.ToString());

                    //dsDETA.SetValue("U_EXC_SEQP", row, Queries.GetSerieItem(oRS.Fields.Item("Code").Value.ToString()));
                    //InstanciateCombo((ComboBox)oMatrix.GetCellSpecific("Col_4", 0), Queries.GetSerieItem(oRS.Fields.Item("Code").Value.ToString()));
                    //while (!oRS.EoF)
                    //{

                    //    oRS.MoveNext();
                    //}
                }
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                LiberarObjetoGenerico(oRS);

            }
        }

        private int GetAbsEntrySerie(string ItemCode, string IMEI)
        {
            SAPbobsCOM.Recordset oRS = null;
            try
            {
                int serie = 0;
                oRS = (SAPbobsCOM.Recordset)Conexion.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRS.DoQuery(Queries.GetAbsEntrySerie(ItemCode, IMEI));

                if (oRS.RecordCount > 0)
                    serie = Convert.ToInt32(oRS.Fields.Item("AbsEntry").Value.ToString());
                return serie;
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                LiberarObjetoGenerico(oRS);

            }
        }

        private string GetInternalNumberSerie(string ItemCode, string IMEI)
        {
            SAPbobsCOM.Recordset oRS = null;
            try
            {
                string serie = string.Empty;
                oRS = (SAPbobsCOM.Recordset)Conexion.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRS.DoQuery(Queries.GetInternalNumberSerie(ItemCode, IMEI));

                if (oRS.RecordCount > 0)
                    serie = oRS.Fields.Item("DistNumber").Value.ToString();
                return serie;
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                LiberarObjetoGenerico(oRS);

            }
        }

        private void SetDetalleSerItem(string ItemCode, string Serie, int row)
        {
            SAPbobsCOM.Recordset oRS = null;
            try
            {
                oRS = (SAPbobsCOM.Recordset)Conexion.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRS.DoQuery(Queries.GetDetalleDesensamble(ItemCode, Serie));

                if (oRS.RecordCount > 0)
                {
                    //dsDETA.SetValue("U_EXC_CEQP", row, oRS.Fields.Item("U_EXC_CEQP").Value.ToString());
                    //dsDETA.SetValue("U_EXC_DEQP", row, oRS.Fields.Item("U_EXC_DEQP").Value.ToString());
                    //dsDETA.SetValue("U_EXC_SEQP", row, oRS.Fields.Item("U_EXC_SEQP").Value.ToString());
                    dsDETA.SetValue("U_EXC_MARC", row, oRS.Fields.Item("U_EXC_MARCA").Value.ToString());
                    dsDETA.SetValue("U_EXC_MODE", row, oRS.Fields.Item("U_EXC_MODELO").Value.ToString());
                    //dsDETA.SetValue("U_EXC_CCHI", row, oRS.Fields.Item("U_EXC_CCHI").Value.ToString());
                    //dsDETA.SetValue("U_EXC_DCHI", row, oRS.Fields.Item("U_EXC_DCHI").Value.ToString());
                    //dsDETA.SetValue("U_EXC_SCHI", row, oRS.Fields.Item("U_EXC_SCHI").Value.ToString());
                    dsDETA.SetValue("U_EXC_IMEI", row, oRS.Fields.Item("U_EXC_IMEI").Value.ToString());
                    dsDETA.SetValue("U_EXC_SEQP", row, oRS.Fields.Item("U_EXC_IMEI").Value.ToString());
                    dsDETA.SetValue("U_EXC_OPER", row, oRS.Fields.Item("U_EXC_OPERAD").Value.ToString());
                    dsDETA.SetValue("U_EXC_ALMI", row, oRS.Fields.Item("U_EXC_ALMI").Value.ToString());
                    dsDETA.SetValue("U_EXC_SCHI", row, oRS.Fields.Item("U_EXC_SIMCARD").Value.ToString());

                    //NUEVOS CAMPOS

                    dsDETA.SetValue("U_EXC_FWAR", row, oRS.Fields.Item("U_EXC_FIRMW").Value.ToString());
                    dsDETA.SetValue("U_EXC_DAPN", row, oRS.Fields.Item("U_EXC_APN").Value.ToString());
                    dsDETA.SetValue("U_EXC_TIIP", row, oRS.Fields.Item("U_EXC_TIPIP").Value.ToString());
                    dsDETA.SetValue("U_EXC_NRIP", row, oRS.Fields.Item("U_EXC_IP").Value.ToString());
                    dsDETA.SetValue("U_EXC_SIMC", row, oRS.Fields.Item("U_EXC_SIMCARD").Value.ToString());
                    dsDETA.SetValue("U_EXC_LINE", row, oRS.Fields.Item("U_EXC_LINEA").Value.ToString());
                    dsDETA.SetValue("U_EXC_FOTA", row, oRS.Fields.Item("U_EXC_FOTA").Value.ToString());
                    dsDETA.SetValue("U_EXC_PQDA", row, oRS.Fields.Item("U_EXC_PAQDATOS").Value.ToString());
                    dsDETA.SetValue("U_EXC_PROP", row, oRS.Fields.Item("U_EXC_PRODPOR").Value.ToString());

                    //Nuevo 20230201
                    dsDETA.SetValue("U_EXC_LTEL", row, oRS.Fields.Item("U_EXC_LINTEL").Value.ToString());
                    dsDETA.SetValue("U_EXC_PROT", row, oRS.Fields.Item("U_EXC_PROTC").Value.ToString());
                    dsDETA.SetValue("U_EXC_NRIP", row, oRS.Fields.Item("U_EXC_IP").Value.ToString());

                    //Nuevo 20230201
                    dsDETA.SetValue("U_EXC_CACC1", row, oRS.Fields.Item("U_EXC_CACC1").Value.ToString());
                    dsDETA.SetValue("U_EXC_DACC1", row, oRS.Fields.Item("U_EXC_DACC1").Value.ToString());
                    dsDETA.SetValue("U_EXC_CACC2", row, oRS.Fields.Item("U_EXC_CACC2").Value.ToString());
                    dsDETA.SetValue("U_EXC_DACC2", row, oRS.Fields.Item("U_EXC_DACC2").Value.ToString());

                    oMatrix.AutoResizeColumns();
                }
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                LiberarObjetoGenerico(oRS);
            }
        }
        private void getNextFormCode(string itemcode)
        {
            dsHEAD.SetValue("Code", 0, itemcode + "-" + GenericQuery(Queries.GetNextFormCode(itemcode)));
        }
        private void UpdateChoose(SAPbouiCOM.DBDataSource ds, SAPbouiCOM.ItemEvent oEvent, SAPbouiCOM.DataTable oDataTable)
        {
            oMatrix = (Matrix)mForm.Items.Item(oEvent.ItemUID).Specific;
            //oMatrix.FlushToDataSource();
            switch (oEvent.ColUID)
            {
                case "Col_Maq":
                    var tes = oDataTable.GetValue("ResCode", 0).ToString().Trim();
                    var te2s = oDataTable.GetValue("U_EXP_VELMAQ", 0);
                    ds.SetValue("U_EXP_RECMAQ", oEvent.Row - 1, oDataTable.GetValue("ResCode", 0).ToString().Trim());
                    ds.SetValue("U_EXP_VELMAQ", oEvent.Row - 1, oDataTable.GetValue("U_EXP_VELMAQ", 0).ToString());
                    break;
                case "Col_MP":
                    var tes1 = oDataTable.GetValue("ItemCode", 0).ToString().Trim();
                    ds.SetValue("U_EXP_MPRIMA", oEvent.Row - 1, oDataTable.GetValue("ItemCode", 0).ToString().Trim());
                    break;
                default:
                    break;
            }
            //oMatrix.LoadFromDataSource();
        }

        public bool HandleFormDataEvents(SAPbouiCOM.BusinessObjectInfo oBusinessObjectInfo)
        {
            switch (oBusinessObjectInfo.EventType)
            {
                case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                    if (mForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && !oBusinessObjectInfo.BeforeAction && oBusinessObjectInfo.ActionSuccess)
                    {

                    }

                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                    if (mForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE && !oBusinessObjectInfo.BeforeAction && oBusinessObjectInfo.ActionSuccess)
                    {
                        WhenDataLoad(oBusinessObjectInfo);
                    }
                    break;
                default:
                    break;
            }
            return true;
        }

        public bool HandleMenuDataEvents(SAPbouiCOM.MenuEvent menuEvent)
        {
            var result = true;
            try
            {
                if (menuEvent.BeforeAction)
                {
                    switch (menuEvent.MenuUID)
                    {
                        case Constants.Menu_EliminarLinea:
                            DeleteRow(RowItemRightClick, ItemUIDRightClick);
                            break;
                        case Constants.Menu_AgregarLinea:
                            AddRow(ItemUIDRightClick);
                            break;
                    }
                }
                else
                {
                    switch (menuEvent.MenuUID)
                    {
                        case Constants.Menu_Crear:
                            dsHEAD.SetValue("U_EXC_ESTA", 0, "O");
                            break;
                        case Constants.Menu_Buscar:
                            break;
                        case Constants.Registro_Datos_Siguiente:
                        case Constants.Registro_Datos_Anterior:
                        case Constants.Primer_Registro_Datos:
                        case Constants.Ultimo_Registro_Datos:

                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                StatusMessageError("HandleMenuDataEvents > " + ex.Message);
            }
            return result;
        }

        #region _EVENTS_RIGHTCLICK
        public bool HandleRightClickEvent(SAPbouiCOM.ContextMenuInfo menuInfo)
        {
            var result = true;
            SAPbouiCOM.MenuItem oMenuItem;
            SAPbouiCOM.Menus oMenus;
            if (menuInfo.BeforeAction)
            {
                try
                {
                    if (mForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || mForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        || mForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                    {
                        SAPbouiCOM.MenuCreationParams oCreationPackage = null;
                        oCreationPackage = (MenuCreationParams)Conexion.application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                        oMenuItem = Conexion.application.Menus.Item(Constants.Menu_Context);
                        oMenus = oMenuItem.SubMenus;

                        ItemUIDRightClick = menuInfo.ItemUID;
                        RowItemRightClick = menuInfo.Row;
                        if (ItemUIDRightClick.StartsWith("mtx"))
                        {
                            if (menuInfo.Row > 0 && !oMenus.Exists(Constants.Menu_AgregarLinea))
                            {
                                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                                oCreationPackage.UniqueID = Constants.Menu_AgregarLinea;
                                oCreationPackage.String = Constants.Menu_AgregarLineaDescripcion;
                                oCreationPackage.Position = 100;
                                oCreationPackage.Enabled = true;
                                oMenus.AddEx(oCreationPackage);
                            }
                            if (menuInfo.Row > 0 && !oMenus.Exists(Constants.Menu_EliminarLinea))
                            {
                                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                                oCreationPackage.UniqueID = Constants.Menu_EliminarLinea;
                                oCreationPackage.String = Constants.Menu_EliminarLineaDescripcion;
                                oCreationPackage.Position = 101;
                                oCreationPackage.Enabled = true;
                                oMenus.AddEx(oCreationPackage);
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    StatusMessageError("HandleRightClickEvent > BeforeAction > " + e.Message);
                }
            }
            else if (!menuInfo.BeforeAction)
            {
                try
                {
                    if (mForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || mForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                    {
                        if (menuInfo.Row > 0)
                            Conexion.application.Menus.RemoveEx(Constants.Menu_EliminarLinea);
                    }
                }
                catch (Exception e)
                {
                    StatusMessageError("HandleRightClickEvent > NotBeforeAction > " + e.Message);
                }
            }
            return result;
        }
        #endregion

        #region _METODOS_PROPIOS


        //private void InstanciateComboRuta(SAPbouiCOM.ComboBox ComboBox, bool noExt = false)
        //{
        //    while (ComboBox.ValidValues.Count != 0)
        //    {
        //        ComboBox.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
        //    }

        //    for (int i = 0; i < dsDETA.Size; i++)
        //    {
        //        if (!noExt || (noExt && dsDETA.GetValue("U_EXP_CODRUT", i).ToString() != Constants.RT_EXTRUS))
        //            ComboBox.ValidValues.Add(dsDETA.GetValue("LineId", i).ToString(), dsDETA.GetValue("U_EXP_NOMRUT", i).ToString());
        //    }
        //    ComboBox.Item.Enabled = true;
        //}

        private string GetRutaDesc(string code, int tipoRef, string TipoRuta, string ruta = "")
        {
            int rutaLinea = string.IsNullOrEmpty(ruta) ? 0 : int.Parse(ruta);
            if (string.IsNullOrEmpty(code)) return "";
            if (int.Parse(code) == rutaLinea) return "";
            switch (tipoRef)
            {
                case Constants.RutaOrigen:
                    if (int.Parse(code) > rutaLinea)
                    {
                        StatusMessageError(string.Format("No se puede poner una ruta posterior", TipoRuta, code));
                        return "";
                    }
                    break;
                case Constants.RutaDestino:
                    if (int.Parse(code) < rutaLinea)
                    {
                        StatusMessageError(string.Format("No se puede poner una ruta anterior", TipoRuta, code));
                        return "";
                    }
                    break;
                default:
                    break;
            }
            string val = "";
            for (int i = 0; i < dsDETA.Size; i++)
            {
                switch (tipoRef)
                {
                    case Constants.RutaBase:
                        if (dsDETA.GetValue("LineId", i).Equals(code) && (dsDETA.GetValue("U_EXP_CODRUT", i).Trim().Equals(TipoRuta.Trim()))) val = dsDETA.GetValue("U_EXP_NOMRUT", i);
                        break;
                    default:
                        if (dsDETA.GetValue("LineId", i).Equals(code)) val = dsDETA.GetValue("U_EXP_NOMRUT", i);
                        break;
                }
            }
            if (string.IsNullOrEmpty(val)) StatusMessageError(string.Format("No hay ruta definida para {0} con código {1}", TipoRuta, code));
            return val;
        }
        private void DeleteRow(int row, string ItemUID)
        {
            try
            {
                mForm.Freeze(true);
                if (dsHEAD.GetValue("U_EXC_ESTA", 0) == "O")
                {
                    oMatrix.FlushToDataSource();
                    dsDETA.RemoveRecord(row - 1);
                    if (mForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        mForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;

                    oMatrix.LoadFromDataSource();
                }
                else
                {
                    StatusMessageWarning("Solo puede eliminar filas en estado abierto.");
                }
            }
            catch (Exception ex)
            {
                StatusMessageError("DeleteRow() > " + ex.Message);
            }
            finally { mForm.Freeze(false); }
        }
        private void AddRow(string ItemUID)
        {
            try
            {
                mForm.Freeze(true);
                oMatrix.FlushToDataSource();
                if (oMatrix.RowCount == 0) dsDETA.Clear();
                dsDETA.InsertRecord(dsDETA.Size);
                dsDETA.SetValue("LineId", dsDETA.Size - 1, dsDETA.Size.ToString());
                oMatrix.LoadFromDataSource();
                oMatrix.AutoResizeColumns();
                ((SAPbouiCOM.EditText)oMatrix.Columns.Item("#").Cells.Item(dsDETA.Size).Specific).Value = dsDETA.Size.ToString();
            }
            catch (Exception ex)
            {
                StatusMessageError("AddRow() > " + ex.Message);
            }
            finally { mForm.Freeze(false); }
        }

        private void CargarArchivo()
        {
            FolderFileDialog openFileDialog;
            SAPbouiCOM.ProgressBar oProgreso = default(SAPbouiCOM.ProgressBar);

            try
            {
                if (mForm.Mode == BoFormMode.fm_ADD_MODE)
                {

                    string Archivo = "";
                    openFileDialog = new FolderFileDialog();
                    Archivo = openFileDialog.FindFile();

                    if (dsDETA.Size > 0)
                    {
                        oMatrix.FlushToDataSource();
                        dsDETA.Clear();
                        oMatrix.LoadFromDataSource();
                    }

                    if (!string.IsNullOrEmpty(Archivo))
                    {
                        StatusMessageInfo("Leyendo archivo excel, por favor espere...");
                        //Leer excel
                        var Excel = new SLDocument(Archivo);
                        string firstSheetName = Excel.GetSheetNames()[0];
                        Excel.SelectWorksheet(firstSheetName);
                        int lastRow = Excel.GetWorksheetStatistics().EndRowIndex;
                        int lastColumn = Excel.GetWorksheetStatistics().EndColumnIndex;
                        System.Data.DataTable dataTable = new System.Data.DataTable();

                        var columnNames = Enumerable.Range(1, lastColumn)
                                           .Select(col => Excel.GetCellValueAsString(1, col))
                                           //.Where(col => !string.IsNullOrWhiteSpace(col))
                                           .ToList();
                        columnNames.ForEach(col => dataTable.Columns.Add(col.Trim()));

                        int articuloIndex = Enumerable.Range(1, lastColumn)
                               .Select(col => Excel.GetCellValueAsString(1, col))
                               .ToList()
                               .FindIndex(col => col.Equals("Cod. Artículo", StringComparison.OrdinalIgnoreCase));

                        var rows = Enumerable.Range(2, lastRow - 1) // Filas de datos
                                      .Select(row => Enumerable.Range(1, lastColumn)
                                                               .Select(col => Excel.GetCellValueAsString(row, col))
                                                               .ToArray())
                                      .Where(row => !string.IsNullOrWhiteSpace(row[articuloIndex]))
                                      .ToList();
                        rows.ForEach(row => dataTable.Rows.Add(row));

                        mForm.Freeze(true);

                        oProgreso = Conexion.application.StatusBar.CreateProgressBar("Cargando", dataTable.Rows.Count, false);
                        oProgreso.Value = 0;

                        foreach (DataRow row in dataTable.Rows)
                        {
                            oMatrix.AddRow();
                            ((EditText)oMatrix.Columns.Item("#").Cells.Item(oMatrix.VisualRowCount).Specific).Value = oMatrix.VisualRowCount.ToString();
                            ((EditText)oMatrix.Columns.Item("Col_0").Cells.Item(oMatrix.VisualRowCount).Specific).Value = row["Cod. Artículo"].ToString();
                            ((EditText)oMatrix.Columns.Item("Col_15").Cells.Item(oMatrix.VisualRowCount).Specific).Value = row["OBC"].ToString();
                            ((EditText)oMatrix.Columns.Item("Col_7").Cells.Item(oMatrix.VisualRowCount).Specific).Value = row["CHIP"].ToString();
                            ((EditText)oMatrix.Columns.Item("Col_14").Cells.Item(oMatrix.VisualRowCount).Specific).Value = row["DESEMBALADO POR"].ToString();
                            ((EditText)oMatrix.Columns.Item("Col_12").Cells.Item(oMatrix.VisualRowCount).Specific).Value = row["ALM - INGRESO"].ToString();
                            ((EditText)oMatrix.Columns.Item("Col_13").Cells.Item(oMatrix.VisualRowCount).Specific).Value = row["ALM - SALIDA"].ToString();
                            ((EditText)oMatrix.Columns.Item("IMEI Chip").Cells.Item(oMatrix.VisualRowCount).Specific).Value = row["IMEI CHIP (SimCard)"].ToString();

                            oProgreso.Value += 1;
                            oProgreso.Text = "Cargando filas " + oProgreso.Value + " de " + dataTable.Rows.Count;
                        }
                        oMatrix.AutoResizeColumns();
                    }
                }
            }
            catch (Exception ex)
            {
                StatusMessageError("Seleccionar excel > " + ex.Message);
                StatusMessageError("Ocurrió un error, es probable que no se hayan copiado todos los datos. Revise el excel");
            }
            finally
            {
                if (oProgreso != null)
                    oProgreso.Stop();
                mForm.Freeze(false);
                GC.Collect();
            }
        }

        private bool AddSpecificMatrix(string matrix, SAPbouiCOM.DBDataSource ds, string aux = "", string codeRut = "")
        {
            try
            {
                //if (!string.IsNullOrEmpty(codeRut) && CheckDataSourceRuta(ds, codeRut, matrix)) return true;
                //mForm.Freeze(true);
                //oMatrix = mForm.Items.Item(matrix).Specific;
                //oMatrix.FlushToDataSource();
                //if (oMatrix.RowCount == 0) ds.Clear();
                //else
                //{
                //    switch (matrix)
                //    {
                //        case MTX_FREXTR:
                //            if (String.IsNullOrEmpty(ds.GetValue("U_EXP_CODRUTDE", 0))) ds.Clear();
                //            break;
                //        default:
                //            if (String.IsNullOrEmpty(ds.GetValue("U_EXP_CODRUT", 0))) ds.Clear();
                //            break;
                //    }
                //}
                //ds.InsertRecord(ds.Size);
                //ds.SetValue("LineId", ds.Size - 1, ds.Size.ToString());
                //switch (matrix)
                //{
                //    case MTX_FORMUL:
                //        ds.SetValue("U_EXP_CODRUT", ds.Size - 1, udRUTA.Value != null ? udRUTA.Value : "");
                //        ds.SetValue("U_EXP_CAPA1", ds.Size - 1, "0");
                //        ds.SetValue("U_EXP_CAPA2", ds.Size - 1, "0");
                //        ds.SetValue("U_EXP_CAPA3", ds.Size - 1, "0");
                //        ds.SetValue("U_EXP_CAPA4", ds.Size - 1, "0");
                //        ds.SetValue("U_EXP_CAPA5", ds.Size - 1, "0");
                //        break;
                //    case MTX_INDUCT:
                //        if (!string.IsNullOrEmpty(codeRut)) ds.SetValue("U_EXP_CODRUT", ds.Size - 1, codeRut);
                //        if (!string.IsNullOrEmpty(aux)) ds.SetValue("U_EXP_INDCTR", ds.Size - 1, aux);
                //        ds.SetValue("U_EXP_STATUS", ds.Size - 1, "Y");
                //        break;
                //    case MTX_FREXTR:
                //        if (!string.IsNullOrEmpty(codeRut)) ds.SetValue("U_EXP_CODRUTDE", ds.Size - 1, codeRut);
                //        if (!string.IsNullOrEmpty(aux)) ds.SetValue("U_EXP_NOMRUTDE", ds.Size - 1, aux);
                //        break;
                //    default:
                //        if (!string.IsNullOrEmpty(codeRut)) ds.SetValue("U_EXP_CODRUT", ds.Size - 1, codeRut);
                //        if (!string.IsNullOrEmpty(aux)) ds.SetValue("U_EXP_NOMRUT", ds.Size - 1, aux);
                //        break;
                //}
                //ds.SetValue("U_Factor", ds.Size - 1, string.Empty);
                //ds.SetValue("U_FactReal", ds.Size - 1, string.Empty);
                oMatrix.LoadFromDataSource();
                oMatrix.AutoResizeColumns();
                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                mForm.Freeze(false);
            }
        }

        private void DeleteSpecificMatrix(string matrix, SAPbouiCOM.DBDataSource ds, int row, bool isClear = false)
        {
            try
            {

                oMatrix = (Matrix)mForm.Items.Item(matrix).Specific;
                oMatrix.FlushToDataSource();
                ds.RemoveRecord(row - 1);

                if (mForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                    mForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                if (matrix.Equals(MTX_MAIN))
                    for (int i = row - 1; i < ds.Size; i++)
                    {
                        ds.SetValue("LineId", i, (i + 1).ToString());
                    }
                oMatrix.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        #endregion

        public string getFormUID()
        {
            if (mForm != null)
                return mForm.UniqueID;
            else
                return null;
        }
    }
}