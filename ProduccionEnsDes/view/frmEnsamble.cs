using AddonProduccionEnsDes.commons;
using AddonProduccionEnsDes.conexion;
using AddonProduccionEnsDes.data_schema;
using DocumentFormat.OpenXml.Drawing;
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
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace AddonProduccionEnsDes.view
{
    public class frmEnsamble : FormCommon, commons.IForm
    {
        #region variables
        private SAPbouiCOM.Form mForm;
        private SAPbouiCOM.DBDataSource dsHEAD;
        private SAPbouiCOM.DBDataSource dsDETA;
        private SAPbouiCOM.Matrix oMatrix;
        private SAPbouiCOM.Item oEdtDate, oEdtWhs;
        private SAPbouiCOM.Item btnOW, btnNS, btnAdd, btnSE, btnCE;

        //CONST PARA LAYOUT
        public const string TYPENAME = "Formulacion";
        public const string ADDONNAME = "ListaMateriales";

        private const string EDT_DATE = "edtDate", EDT_WHS = "edtWHS"; //EditTexts Porc
        private const string BTN_OK = "1", BTN_ORDFAB = "btnOrd", BTN_ENSAMBLAR = "btnExe", BTN_ADD = "btnAdd", BTN_SERIE = "btnSerie", BTN_FILE = "btnFile";//Buttons
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

        public frmEnsamble(Dictionary<string, commons.IForm> dictionary)
        {
            try
            {
                mForm = CreateForm(Conexion.company, Conexion.application, Properties.Resources.frmEnsamble, FormName.ENSAMBLE);
                if (mForm != null)
                {
                    if (Conexion.application.ClientType == BoClientType.ct_Browser)
                    {
                        SAPbouiCOM.Item btnCancelar = (SAPbouiCOM.Item)mForm.Items.Item("2");

                        SAPbouiCOM.Item oItem2 = mForm.Items.Add("RUTA", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        oItem2.Left = btnCancelar.Left + btnCancelar.Width + (btnCancelar.Width * 3 / 2) + 10;
                        oItem2.Width = mForm.Width - (oItem2.Left + 20);
                        oItem2.Top = btnCancelar.Top;//80;
                        oItem2.Height = btnCancelar.Height - 2;
                        SAPbouiCOM.EditText oEditText2 = ((SAPbouiCOM.EditText)(oItem2.Specific));
                        oEditText2.String = "";

                        SAPbouiCOM.Item oIteml2 = mForm.Items.Add("RUTAL", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                        oIteml2.Left = btnCancelar.Left + btnCancelar.Width + +10;
                        oIteml2.Width = btnCancelar.Width * 3 / 2 - 30;
                        oIteml2.Top = btnCancelar.Top;
                        oIteml2.Height = btnCancelar.Height - 2;
                        SAPbouiCOM.StaticText oStaticText2 = ((SAPbouiCOM.StaticText)(oIteml2.Specific));
                        oStaticText2.Caption = "Ingrese ruta de Excel";
                        oStaticText2.Item.LinkTo = "RUTA";
                    }

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
                if (dsHEAD == null) dsHEAD = mForm.DataSources.DBDataSources.Item($"@{SCEnsamble.TABLE_CABE}");
                dsDETA = mForm.DataSources.DBDataSources.Item($"@{SCEnsamble.TABLE_DET1}");
                oMatrix = (SAPbouiCOM.Matrix)mForm.Items.Item(MTX_MAIN).Specific;
                oEdtDate = (SAPbouiCOM.Item)mForm.Items.Item(EDT_DATE);
                oEdtWhs = (SAPbouiCOM.Item)mForm.Items.Item(EDT_WHS);
                btnOW = (SAPbouiCOM.Item)mForm.Items.Item(BTN_ORDFAB);
                btnNS = (SAPbouiCOM.Item)mForm.Items.Item(BTN_ENSAMBLAR);
                btnAdd = (SAPbouiCOM.Item)mForm.Items.Item(BTN_ADD);
                btnCE = (SAPbouiCOM.Item)mForm.Items.Item(BTN_FILE);
                btnSE = (SAPbouiCOM.Item)mForm.Items.Item(BTN_SERIE);
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

                oChooseFromList = oChooseFromListCollection.Item("cflLMAT");
                oConditions = oChooseFromList.GetConditions();
                //oCondition = oConditions.Add();
                //oCondition.Alias = "TreeType";
                //oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                //oCondition.CondVal = "P";
                oChooseFromList.SetConditions(oConditions);

                oChooseFromList = oChooseFromListCollection.Item("cflCHIP");
                oConditions = oChooseFromList.GetConditions();
                oCondition = oConditions.Add();
                oCondition.Alias = "U_EXC_DISPROD";
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCondition.CondVal = "Y";//SimCards
                oChooseFromList.SetConditions(oConditions);

                oChooseFromList = oChooseFromListCollection.Item("cflICHI");
                oConditions = oChooseFromList.GetConditions();
                oCondition = oConditions.Add();
                oCondition.Alias = "U_EXC_IMEI";
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                oCondition.CondVal = "";//        
                oChooseFromList.SetConditions(oConditions);

                oChooseFromList = oChooseFromListCollection.Item("cflIEQP");
                oConditions = oChooseFromList.GetConditions();
                oCondition = oConditions.Add();
                oCondition.Alias = "U_EXC_IMEI";
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                oCondition.CondVal = "";//            
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
                //InstanciateCombo((ComboBox)oMatrix.GetCellSpecific("Col_9", 0), Queries.GetSerieItem());
                //InstanciateCombo((ComboBox)oMatrix.GetCellSpecific("Col_4", 0), Queries.GetSerieItem());
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
                        //if (!itemEvent.BeforeAction)
                        result = whenChooseFromList(itemEvent);
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
        private bool MatrixValidate(SAPbouiCOM.ItemEvent oEvent)
        {
            bool res = true;
            oMatrix = (Matrix)mForm.Items.Item(oEvent.ItemUID).Specific;



            switch (oEvent.ColUID)
            {
                case "Col_4": //SERIE EQUIPO

                    string serie = ((EditText)oMatrix.GetCellSpecific("Col_4", oEvent.Row)).Value.ToString();

                    if (!string.IsNullOrEmpty(serie) && oEvent.ItemChanged)
                    {
                        oMatrix.FlushToDataSource();

                        res = ValidarDuplicidadSerieEquipo(dsDETA.GetValue("U_EXC_CEQP", oEvent.Row - 1), dsDETA.GetValue("U_EXC_SEQP", oEvent.Row - 1), false);

                        if (res)
                        {
                            SetDetalleSerItem(dsDETA.GetValue("U_EXC_CEQP", oEvent.Row - 1), dsDETA.GetValue("U_EXC_SEQP", oEvent.Row - 1), oEvent.Row - 1, false);
                        }
                        else
                        {
                            dsDETA.SetValue("U_EXC_SEQP", oEvent.Row - 1, "");
                            dsDETA.SetValue("U_EXC_MARC", oEvent.Row - 1, "");
                            dsDETA.SetValue("U_EXC_MODE", oEvent.Row - 1, "");
                            dsDETA.SetValue("U_EXC_IMEIEQP", oEvent.Row - 1, "");
                            StatusMessageWarning("No puede ponerse serie repetida");
                        }

                        oMatrix.LoadFromDataSource();
                    }

                    break;

                case "Col_9": //SERIE CHIP 

                    oMatrix.FlushToDataSource();
                    res = ValidarDuplicidadSerieEquipo(dsDETA.GetValue("U_EXC_CCHI", oEvent.Row - 1), dsDETA.GetValue("U_EXC_SCHI", oEvent.Row - 1), true);

                    if (res)
                    {
                        SetDetalleSerItem(dsDETA.GetValue("U_EXC_CCHI", oEvent.Row - 1), dsDETA.GetValue("U_EXC_SCHI", oEvent.Row - 1), oEvent.Row - 1);
                    }
                    else
                    {
                        dsDETA.SetValue("U_EXC_SCHI", oEvent.Row - 1, "");
                        dsDETA.SetValue("U_EXC_IMEI", oEvent.Row - 1, "");
                        dsDETA.SetValue("U_EXC_OPER", oEvent.Row - 1, "");
                        StatusMessageWarning("No puede ponerse serie repetida");
                    }

                    oMatrix.LoadFromDataSource();
                    break;
                default:
                    break;


            }

            oMatrix.AutoResizeColumns();

            return true;
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
                            oMatrix.Item.Enabled = true;
                            oEdtDate.Enabled = true;
                            oEdtWhs.Enabled = true;
                            switch (dsHEAD.GetValue("U_EXC_ESTA", 0))
                            {
                                case "O":
                                    actOW = true;
                                    actNS = false;
                                    break;
                                case "P":
                                    actOW = false;
                                    actNS = true;
                                    break;
                                case "F":
                                    oMatrix.Item.Enabled = false;
                                    oEdtDate.Enabled = false;
                                    oEdtWhs.Enabled = false;
                                    actOW = false;
                                    actNS = false;
                                    break;
                                default:
                                    break;
                            }
                            btnAdd.Enabled = false;
                            btnCE.Enabled = false;
                            btnOW.Enabled = actOW;
                            btnNS.Enabled = actNS;
                            btnAdd.Enabled = actNS;
                        }
                        break;
                    case BTN_ORDFAB:
                        if (oEvent.ActionSuccess)
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
                            if (res) res = ProcesoEnsamblar(oEvent);
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

        private bool ProcesoEnsamblar(ItemEvent oEvent)
        {
            bool sta = true;
            string Estado = "F";
            int res;
            List<string> Series = new List<string>();
            try
            {
                if (mForm.Mode != BoFormMode.fm_ADD_MODE && mForm.Mode != BoFormMode.fm_FIND_MODE)
                {
                    if (dsHEAD.GetValue("U_EXC_ESTA", 0) != "P")
                    {
                        StatusMessageWarning("Primero debe generar la orden de fabricación");
                        return false;
                    }
                    StatusMessageWarning("Iniciando proceso de ensamblaje.");

                    //oMatrix.FlushToDataSource();
                    for (int i = 0; i < dsDETA.Size; i++)
                    {
                        if (string.IsNullOrEmpty(dsDETA.GetValue("U_EXC_NSER", i)))
                        {
                            try
                            {
                                string AlmacenOT = GetAlmacenOT(dsDETA.GetValue("U_EXC_ORDT", i));
                                if (dsHEAD.GetValue("U_EXC_ALMA", 0) != AlmacenOT)
                                {
                                    StatusMessageWarning("El almacén de la Orden de Trabajo " + dsDETA.GetValue("U_EXC_ORDT", i) + " no coincide con el del ensamble, por favor rectifique.");
                                    return false;
                                }

                                Conexion.company.StartTransaction();
                                sta = ActualizarOrdProd(i, BoProductionOrderStatusEnum.boposReleased);
                                if (sta) sta = EmitirProducto(i);
                                if (sta) sta = GetSerieTerminado(i, ref Series);
                                if (sta) sta = EntregarProducto(i, Series);
                                if (sta) sta = ActualizarOrdProd(i, BoProductionOrderStatusEnum.boposClosed);

                                if (sta && Conexion.company.InTransaction)
                                {
                                    //dsHEAD.SetValue("U_EXC_ESTA", 0, "F");
                                    GenericQuery(Queries.UpdateSerieUDO(dsHEAD.GetValue("DocEntry", 0), dsDETA.GetValue("U_EXC_ORDT", i), Series[i]));
                                    GenericQuery(Queries.UpdateSerieTerminado(dsDETA.GetValue("U_EXC_MARC", i), dsDETA.GetValue("U_EXC_MODE", i)));
                                    Conexion.company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                                }
                                //if (Conexion.company.InTransaction) Conexion.company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                            }
                            catch (Exception ex)
                            {
                                if (Series.Count == i + 1)
                                    Series[i] = "";

                                if (Conexion.company.InTransaction) Conexion.company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                sta = false;
                                Estado = "P";

                                int errorCode; string errorMessage;
                                Conexion.company.GetLastError(out errorCode, out errorMessage);

                                throw new Exception(string.Format("Falló en ensamblar: {0}", errorMessage));
                            }
                        }
                        else
                        {
                            Series.Add(dsDETA.GetValue("U_EXC_NSER", i));
                        }
                        StatusMessageWarning("Generando ensamblajes " + (i + 1) + " de " + dsDETA.Size);
                    }
                    //oMatrix.LoadFromDataSource();

                    //if (mForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                    //    mForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    //mForm.Items.Item(BTN_OK).Click();

                }
                else
                {
                    StatusMessageError("El registro debe estar creado para procesar el ensamble");
                    return false;
                }
            }
            catch (Exception ex)
            {
                StatusMessageError(ex.Message);

                //oMatrix.LoadFromDataSource();
                //if (mForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                //    mForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                //mForm.Items.Item(BTN_OK).Click();
                StatusMessageWarning("Generación de ensamblajes culminó con errores.");
            }

            if (Series.Count > 0)
            {
                try
                {
                    SAPbobsCOM.CompanyService oCmpSrv = Conexion.company.GetCompanyService();
                    SAPbobsCOM.GeneralService oGeneralService = oCmpSrv.GetGeneralService("EXC_ENSA");
                    SAPbobsCOM.GeneralDataParams oGeneralParams = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);

                    oGeneralParams.SetProperty("DocEntry", dsHEAD.GetValue("DocEntry", 0));
                    SAPbobsCOM.GeneralData oGeneralData = oGeneralService.GetByParams(oGeneralParams);
                    oGeneralData.SetProperty("U_EXC_ESTA", Estado);

                    SAPbobsCOM.GeneralDataCollection oChildren = oGeneralData.Child("EXC_ENS1");
                    //for (int i = 0; i < Series.Count; i++)
                    //{
                    //    SAPbobsCOM.GeneralData oChild = oChildren.Item(i);
                    //    oChild.SetProperty("U_EXC_NSER", Series[i]);
                    //}
                    oGeneralService.Update(oGeneralData);
                    if (Conexion.application.Menus.Item("1304").Enabled)
                        Conexion.application.Menus.Item("1304").Activate();
                    StatusMessageWarning("Generación de ensamblajes culminó satisfactoriamente.");

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCmpSrv);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGeneralService);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGeneralParams);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGeneralData);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oChildren);
                }
                catch (Exception ex)
                {

                    throw;
                }
            }
            return sta;
        }

        private bool GetSerieTerminado(int row, ref List<string> series)
        {
            bool sta = true;
            string serie = GenericQuery(Queries.GetSerieTerminado(dsDETA.GetValue("U_EXC_MARC", row), dsDETA.GetValue("U_EXC_MODE", row)));
            if (string.IsNullOrEmpty(serie))
            {
                series.Add("");
                throw new Exception(string.Format("No se tiene serie correlativo para marca {0} y modelo {1} ", dsDETA.GetValue("U_EXC_MARC", row), dsDETA.GetValue("U_EXC_MODE", row)));
                //if (Conexion.company.InTransaction) Conexion.company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                //return false;
            }
            else
            {
                series.Add(serie);
                //dsDETA.SetValue("U_EXC_NSER", row, serie);
                //oMatrix.Columns.Item("Col_15").Cells.Item(row + 1).Specific.Value = serie;
            }
            return sta;
        }

        private bool EntregarProducto(int row, List<string> series)
        {
            //ACA MODIFICAR
            Documents oReciboProduccion = null;
            bool sta = true;
            int res;
            try
            {
                oReciboProduccion = (Documents)Conexion.company.GetBusinessObject(BoObjectTypes.oInventoryGenEntry);
                oReciboProduccion.DocDate = DateTime.ParseExact((dsHEAD.GetValue("U_EXC_FEPR", 0)), "yyyyMMdd", CultureInfo.InvariantCulture);

                oReciboProduccion.UserFields.Fields.Item("U_EXX_TIPOOPER").Value = AddonProduccionEnsDes.Properties.Resources.TOperacionEntrada;

                oReciboProduccion.Lines.BaseType = (int)BoObjectTypes.oProductionOrders;
                oReciboProduccion.Lines.BaseEntry = int.Parse(dsDETA.GetValue("U_EXC_ORDT", row));
                oReciboProduccion.Lines.Quantity = 1;

                //oReciboProduccion.Lines.SerialNumbers.InternalSerialNumber = dsDETA.GetValue("U_EXC_NSER", row);
                oReciboProduccion.Lines.SerialNumbers.InternalSerialNumber = series[row];
                oReciboProduccion.Lines.SerialNumbers.ManufacturerSerialNumber = dsDETA.GetValue("U_EXC_IMEIEQP", row);
                oReciboProduccion.Lines.SerialNumbers.ManufactureDate = DateTime.ParseExact((dsHEAD.GetValue("U_EXC_FEPR", 0)), "yyyyMMdd", CultureInfo.InvariantCulture);

                oReciboProduccion.Lines.SerialNumbers.UserFields.Fields.Item("U_EXC_VERSION").Value = dsDETA.GetValue("U_EXC_VERS", row);
                oReciboProduccion.Lines.SerialNumbers.UserFields.Fields.Item("U_EXC_ARCONF").Value = dsDETA.GetValue("U_EXC_ARCH", row);
                oReciboProduccion.Lines.SerialNumbers.UserFields.Fields.Item("U_EXC_PRODPOR").Value = dsDETA.GetValue("U_EXC_PROP", row);
                oReciboProduccion.Lines.SerialNumbers.UserFields.Fields.Item("U_EXC_MARCA").Value = dsDETA.GetValue("U_EXC_MARC", row);
                oReciboProduccion.Lines.SerialNumbers.UserFields.Fields.Item("U_EXC_MODELO").Value = dsDETA.GetValue("U_EXC_MODE", row);
                oReciboProduccion.Lines.SerialNumbers.UserFields.Fields.Item("U_EXC_IMEI").Value = dsDETA.GetValue("U_EXC_IMEIEQP", row);
                oReciboProduccion.Lines.SerialNumbers.UserFields.Fields.Item("U_EXC_OPERAD").Value = dsDETA.GetValue("U_EXC_OPER", row);

                oReciboProduccion.Lines.SerialNumbers.UserFields.Fields.Item("U_EXC_FIRMW").Value = dsDETA.GetValue("U_EXC_VERS", row);
                oReciboProduccion.Lines.SerialNumbers.UserFields.Fields.Item("U_EXC_FOTA").Value = dsDETA.GetValue("U_EXC_FOTA", row);
                oReciboProduccion.Lines.SerialNumbers.UserFields.Fields.Item("U_EXC_TIPIP").Value = dsDETA.GetValue("U_EXC_TIIP", row);
                oReciboProduccion.Lines.SerialNumbers.UserFields.Fields.Item("U_EXC_IP").Value = dsDETA.GetValue("U_EXC_NRIP", row);
                oReciboProduccion.Lines.SerialNumbers.UserFields.Fields.Item("U_EXC_PAQDATOS").Value = dsDETA.GetValue("U_EXC_PQDA", row);
                oReciboProduccion.Lines.SerialNumbers.UserFields.Fields.Item("U_EXC_LINEA").Value = dsDETA.GetValue("U_EXC_LINE", row);
                oReciboProduccion.Lines.SerialNumbers.UserFields.Fields.Item("U_EXC_LINTEL").Value = dsDETA.GetValue("U_EXC_LTEL", row);
                oReciboProduccion.Lines.SerialNumbers.UserFields.Fields.Item("U_EXC_APN").Value = dsDETA.GetValue("U_EXC_DAPN", row);

                oReciboProduccion.Lines.SerialNumbers.UserFields.Fields.Item("U_EXC_SERVI").Value = dsDETA.GetValue("U_EXC_SERV", row);
                oReciboProduccion.Lines.SerialNumbers.UserFields.Fields.Item("U_EXC_SIMCARD").Value = dsDETA.GetValue("U_EXC_SIMC", row);
                oReciboProduccion.Lines.SerialNumbers.UserFields.Fields.Item("U_EXC_PROTC").Value = dsDETA.GetValue("U_EXC_PROT", row);
                oReciboProduccion.Lines.SerialNumbers.UserFields.Fields.Item("U_EXC_DISTNUM").Value = dsDETA.GetValue("U_EXC_NSLO", row);
                oReciboProduccion.Lines.SerialNumbers.UserFields.Fields.Item("U_EXC_CACC1").Value = dsDETA.GetValue("U_EXC_CACC1", row);
                oReciboProduccion.Lines.SerialNumbers.UserFields.Fields.Item("U_EXC_CACC2").Value = dsDETA.GetValue("U_EXC_CACC2", row);

                //if (!string.IsNullOrEmpty(dsDETA.GetValue("U_EXC_CACC1", row))) oReciboProduccion.Lines.SerialNumbers.UserFields.Fields.Item("U_EXC_CACC1").Value = dsDETA.GetValue("U_EXC_CACC1", row);
                //else oReciboProduccion.Lines.SerialNumbers.UserFields.Fields.Item("U_EXC_CACC1").Value = "";
                //if (!string.IsNullOrEmpty(dsDETA.GetValue("U_EXC_CACC2", row))) oReciboProduccion.Lines.SerialNumbers.UserFields.Fields.Item("U_EXC_CACC2").Value = dsDETA.GetValue("U_EXC_CACC2", row);
                //else oReciboProduccion.Lines.SerialNumbers.UserFields.Fields.Item("U_EXC_CACC2").Value = "";

                oReciboProduccion.Lines.SerialNumbers.Quantity = 1;
                res = oReciboProduccion.Add();

                if (res != 0)
                {
                    sta = false;
                    if (Conexion.company.InTransaction) Conexion.company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    //throw new Exception(string.Format("Falló la fabricación:{0}", Conexion.company.GetLastErrorDescription()));
                    throw new Exception(Conexion.company.GetLastErrorDescription());
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                LiberarObjetoGenerico(oReciboProduccion);
            }

            return sta;
        }

        private bool EmitirProducto(int row)
        {
            Documents oEmisionProduccion = null;
            bool sta = true;
            int res;
            try
            {
                oEmisionProduccion = (Documents)Conexion.company.GetBusinessObject(BoObjectTypes.oInventoryGenExit);
                oEmisionProduccion.DocDate = DateTime.ParseExact((dsHEAD.GetValue("U_EXC_FEPR", 0)), "yyyyMMdd", CultureInfo.InvariantCulture);
                oEmisionProduccion.UserFields.Fields.Item("U_EXX_TIPOOPER").Value = AddonProduccionEnsDes.Properties.Resources.TOperacionSalida;

                int index = 0;
                oEmisionProduccion.Lines.SetCurrentLine(index);
                oEmisionProduccion.Lines.BaseType = (int)BoObjectTypes.oProductionOrders;
                oEmisionProduccion.Lines.BaseEntry = int.Parse(dsDETA.GetValue("U_EXC_ORDT", row));
                oEmisionProduccion.Lines.BaseLine = index;
                oEmisionProduccion.Lines.Quantity = 1;
                oEmisionProduccion.Lines.SerialNumbers.SystemSerialNumber = int.Parse(dsDETA.GetValue("U_EXC_SEQP", row));
                oEmisionProduccion.Lines.SerialNumbers.Quantity = 1;
                oEmisionProduccion.Lines.WarehouseCode = dsHEAD.GetValue("U_EXC_ALMA", 0);
                index++;

                oEmisionProduccion.Lines.Add();
                oEmisionProduccion.Lines.SetCurrentLine(index);
                oEmisionProduccion.Lines.BaseType = (int)BoObjectTypes.oProductionOrders;
                oEmisionProduccion.Lines.BaseEntry = int.Parse(dsDETA.GetValue("U_EXC_ORDT", row));
                oEmisionProduccion.Lines.BaseLine = index;
                oEmisionProduccion.Lines.Quantity = 1;
                oEmisionProduccion.Lines.SerialNumbers.SystemSerialNumber = int.Parse(dsDETA.GetValue("U_EXC_SCHI", row));
                oEmisionProduccion.Lines.SerialNumbers.Quantity = 1;
                oEmisionProduccion.Lines.WarehouseCode = AddonProduccionEnsDes.Properties.Resources.AlmSalida;
                index++;

                //ACCESORIOS
                if (!string.IsNullOrEmpty(dsDETA.GetValue("U_EXC_CACC1", row)))
                {
                    oEmisionProduccion.Lines.Add();
                    oEmisionProduccion.Lines.SetCurrentLine(index);
                    oEmisionProduccion.Lines.BaseType = (int)BoObjectTypes.oProductionOrders;
                    oEmisionProduccion.Lines.BaseEntry = int.Parse(dsDETA.GetValue("U_EXC_ORDT", row));
                    oEmisionProduccion.Lines.BaseLine = index;
                    oEmisionProduccion.Lines.Quantity = 1;
                    oEmisionProduccion.Lines.WarehouseCode = AddonProduccionEnsDes.Properties.Resources.AlmSalida;
                    index++;
                }
                if (!string.IsNullOrEmpty(dsDETA.GetValue("U_EXC_CACC2", row)))
                {
                    oEmisionProduccion.Lines.Add();
                    oEmisionProduccion.Lines.SetCurrentLine(index);
                    oEmisionProduccion.Lines.BaseType = (int)BoObjectTypes.oProductionOrders;
                    oEmisionProduccion.Lines.BaseEntry = int.Parse(dsDETA.GetValue("U_EXC_ORDT", row));
                    oEmisionProduccion.Lines.BaseLine = index;
                    oEmisionProduccion.Lines.Quantity = 1;
                    oEmisionProduccion.Lines.WarehouseCode = AddonProduccionEnsDes.Properties.Resources.AlmSalida;
                }


                res = oEmisionProduccion.Add();
                if (res != 0)
                {
                    if (Conexion.company.InTransaction) Conexion.company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    sta = false;
                    //throw new Exception(string.Format("Falló la fabricación:{0}", Conexion.company.GetLastErrorDescription()));
                    throw new Exception(Conexion.company.GetLastErrorDescription());
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                LiberarObjetoGenerico(oEmisionProduccion);
            }

            return sta;
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

                if (status == BoProductionOrderStatusEnum.boposClosed)
                    oOrdeProd.ClosingDate = oOrdeProd.StartDate;

                oOrdeProd.ProductionOrderStatus = status;
                res = oOrdeProd.Update();
                if (res != 0)
                {
                    sta = false;
                    if (Conexion.company.InTransaction) Conexion.company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    //throw new Exception(string.Format("Falló la fabricación:{0}", Conexion.company.GetLastErrorDescription()));
                    throw new Exception(Conexion.company.GetLastErrorDescription());
                }
            }
            catch (Exception ex)
            {
                throw ex;
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
                    StatusMessageWarning("Iniciando proceso de generación orden de fabricación.");

                    oMatrix.FlushToDataSource();

                    for (int i = 0; i < dsDETA.Size; i++)
                    {
                        if (string.IsNullOrEmpty(dsDETA.GetValue("U_EXC_ORDT", i)))
                        {
                            oProduction = (ProductionOrders)Conexion.company.GetBusinessObject(BoObjectTypes.oProductionOrders);
                            oProduction.PostingDate = DateTime.ParseExact((dsHEAD.GetValue("U_EXC_FEPR", 0)), "yyyyMMdd", CultureInfo.InvariantCulture);
                            oProduction.DueDate = DateTime.ParseExact((dsHEAD.GetValue("U_EXC_FEPR", 0)), "yyyyMMdd", CultureInfo.InvariantCulture);
                            oProduction.ItemNo = dsDETA.GetValue("U_EXC_CPRO", i);
                            oProduction.PlannedQuantity = 1;
                            oProduction.Warehouse = dsHEAD.GetValue("U_EXC_ALMA", 0);

                            //EQUIPO
                            oProduction.Lines.ItemNo = dsDETA.GetValue("U_EXC_CEQP", i);
                            //oProduction.Lines.Warehouse = GenericQuery(Queries.GetWhsSerie(dsDETA.GetValue("U_EXC_CEQP", i), dsDETA.GetValue("U_EXC_SEQP", i)));
                            oProduction.Lines.Warehouse = dsHEAD.GetValue("U_EXC_ALMA", 0);
                            oProduction.Lines.ProductionOrderIssueType = BoIssueMethod.im_Manual;
                            oProduction.Lines.Add();

                            //CHIP
                            oProduction.Lines.ItemNo = dsDETA.GetValue("U_EXC_CCHI", i);
                            //oProduction.Lines.Warehouse = GenericQuery(Queries.GetWhsSerie(dsDETA.GetValue("U_EXC_CCHI", i), dsDETA.GetValue("U_EXC_SCHI", i)));
                            oProduction.Lines.Warehouse = AddonProduccionEnsDes.Properties.Resources.AlmSalida; //dsHEAD.GetValue("U_EXC_ALMA", 0);
                            oProduction.Lines.ProductionOrderIssueType = BoIssueMethod.im_Manual;
                            oProduction.Lines.Add();

                            //ACCESORIOS
                            if (!string.IsNullOrEmpty(dsDETA.GetValue("U_EXC_CACC1", i)))
                            {
                                oProduction.Lines.ItemNo = dsDETA.GetValue("U_EXC_CACC1", i);
                                oProduction.Lines.Warehouse = AddonProduccionEnsDes.Properties.Resources.AlmSalida; //dsHEAD.GetValue("U_EXC_ALMA", 0);
                                oProduction.Lines.ProductionOrderIssueType = BoIssueMethod.im_Manual;
                                oProduction.Lines.Add();
                            }
                            if (!string.IsNullOrEmpty(dsDETA.GetValue("U_EXC_CACC2", i)))
                            {
                                oProduction.Lines.ItemNo = dsDETA.GetValue("U_EXC_CACC2", i);
                                oProduction.Lines.Warehouse = AddonProduccionEnsDes.Properties.Resources.AlmSalida; //dsHEAD.GetValue("U_EXC_ALMA", 0);
                                oProduction.Lines.ProductionOrderIssueType = BoIssueMethod.im_Manual;
                                oProduction.Lines.Add();
                            }

                            res = oProduction.Add();
                            //oProduction.SaveXML("D:\\produccion.xml");
                            if (res != 0)
                            {
                                StatusMessageError(string.Format("Falló la fabricación:{0}", Conexion.company.GetLastErrorDescription()));
                                sta = false;
                            }
                            else
                            {
                                string DocEntry = Conexion.company.GetNewObjectKey();
                                dsDETA.SetValue("U_EXC_ORDT", i, DocEntry);
                                ((EditText)oMatrix.Columns.Item("Col_16").Cells.Item(i + 1).Specific).Value = DocEntry;
                            }
                        }
                        StatusMessageWarning("Creando órdenes de fabricación " + (i + 1) + " de " + dsDETA.Size);
                    }
                    oMatrix.LoadFromDataSource();
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
                    case "Col_4": //SERIE EQUIPO
                        oMatrix.FlushToDataSource();

                        res = ValidarDuplicidadSerieEquipo(dsDETA.GetValue("U_EXC_CEQP", oEvent.Row - 1), dsDETA.GetValue("U_EXC_SEQP", oEvent.Row - 1), false);
                        if (res)
                        {
                            SetDetalleSerItem(dsDETA.GetValue("U_EXC_CEQP", oEvent.Row - 1), dsDETA.GetValue("U_EXC_SEQP", oEvent.Row - 1), oEvent.Row - 1, false);
                        }
                        else
                        {
                            dsDETA.SetValue("U_EXC_SEQP", oEvent.Row - 1, "");
                            dsDETA.SetValue("U_EXC_MARC", oEvent.Row - 1, "");
                            dsDETA.SetValue("U_EXC_MODE", oEvent.Row - 1, "");
                            dsDETA.SetValue("U_EXC_IMEIEQP", oEvent.Row - 1, "");
                            StatusMessageWarning("No puede ponerse serie repetida");
                        }

                        oMatrix.LoadFromDataSource();

                        break;

                    case "Col_9": //SERIE CHIP

                        oMatrix.FlushToDataSource();
                        res = ValidarDuplicidadSerieEquipo(dsDETA.GetValue("U_EXC_CCHI", oEvent.Row - 1), dsDETA.GetValue("U_EXC_SCHI", oEvent.Row - 1), true);

                        if (res)
                        {
                            SetDetalleSerItem(dsDETA.GetValue("U_EXC_CCHI", oEvent.Row - 1), dsDETA.GetValue("U_EXC_SCHI", oEvent.Row - 1), oEvent.Row - 1);
                        }
                        else
                        {
                            dsDETA.SetValue("U_EXC_SCHI", oEvent.Row - 1, "");
                            dsDETA.SetValue("U_EXC_IMEI", oEvent.Row - 1, "");
                            dsDETA.SetValue("U_EXC_OPER", oEvent.Row - 1, "");
                            StatusMessageWarning("No puede ponerse serie repetida");
                        }

                        oMatrix.LoadFromDataSource();
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

        private bool ValidarDuplicidadSerieEquipo(string ItemCode, string seriesId, bool isChip)
        {
            bool res = true;
            int count = int.Parse(GenericQuery(Queries.CheckPreviousSeries(ItemCode, seriesId, isChip)));


            for (int i = 0; i < dsDETA.Size; i++)
            {
                if (dsDETA.GetValue(isChip ? "U_EXC_SCHI" : "U_EXC_SEQP", i) == seriesId && dsDETA.GetValue(isChip ? "U_EXC_CCHI" : "U_EXC_CEQP", i) == ItemCode) count++;
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
            oMatrix.Item.Enabled = true;
            oEdtDate.Enabled = true;
            oEdtWhs.Enabled = true;
            switch (dsHEAD.GetValue("U_EXC_ESTA", 0))
            {
                case "O":
                    actOW = true;
                    actNS = false;
                    break;
                case "P":
                    actOW = false;
                    actNS = true;
                    break;
                case "F":
                    oMatrix.Item.Enabled = false;
                    oEdtDate.Enabled = false;
                    oEdtWhs.Enabled = false;
                    actOW = false;
                    actNS = false;
                    break;
                default:
                    break;
            }
            btnAdd.Enabled = false;
            btnCE.Enabled = false;
            btnOW.Enabled = actOW;
            btnNS.Enabled = actNS;
            btnAdd.Enabled = actNS;
            return res;
        }
        #endregion
        private bool whenChooseFromList(SAPbouiCOM.ItemEvent oEvent)
        {
            bool result = true;

            try
            {
                SAPbouiCOM.IChooseFromListEvent oChooseFromListEvent = (SAPbouiCOM.IChooseFromListEvent)oEvent;
                SAPbouiCOM.DataTable oDataTable = oChooseFromListEvent.SelectedObjects;

                switch (oEvent.ItemUID)
                {
                    case MTX_MAIN:
                        result = whenMatrixChooseFromList(oEvent);
                        break;
                    default:
                        break;
                }

                //if (oDataTable != null)
                //{

                //}
            }
            catch (Exception e)
            {
                StatusMessageError("whenChooseFromList > " + e.Message);
            }

            return result;
        }
        private bool whenMatrixChooseFromList(SAPbouiCOM.ItemEvent oEvent)
        {
            bool result = true;

            try
            {
                if (oEvent.BeforeAction) //Filtros dinámicos CFL
                {
                    switch (oEvent.ColUID)
                    {
                        case "Col_17": FiltrarIMEIEquipo(oEvent, ref result); break;
                        case "Col_10": FiltrarIMEIChip(oEvent, ref result); break;
                        default:
                            break;
                    }
                }

                if (!oEvent.BeforeAction)
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
                                dsHEAD.SetValue("U_EXC_ALMA", oEvent.Row - 1, oDataTable.GetValue("ToWH", 0).ToString().Trim());
                                SetDetalleProducto(oDataTable.GetValue("Code", 0).ToString().Trim(), oEvent.Row - 1);
                                oMatrix.LoadFromDataSource();
                                break;
                            case "Col_7":
                                oMatrix.FlushToDataSource();
                                dsDETA.SetValue("U_EXC_CCHI", oEvent.Row - 1, oDataTable.GetValue("ItemCode", 0).ToString().Trim());
                                dsDETA.SetValue("U_EXC_DCHI", oEvent.Row - 1, oDataTable.GetValue("ItemName", 0).ToString().Trim());

                                //InstanciateCombo((ComboBox)oMatrix.GetCellSpecific("Col_9", 0), Queries.GetSerieItem(oDataTable.GetValue("ItemCode", 0).Trim()), false);

                                RefreshChooseFromlist("cflICHI", oDataTable.GetValue("ItemCode", 0).ToString().Trim(), true);
                                oMatrix.LoadFromDataSource();
                                break;
                            case "Col_17":
                                oMatrix.FlushToDataSource();
                                dsDETA.SetValue("U_EXC_IMEIEQP", oEvent.Row - 1, oDataTable.GetValue("U_EXC_IMEI", 0).ToString().Trim());
                                dsDETA.SetValue("U_EXC_MARC", oEvent.Row - 1, oDataTable.GetValue("U_EXC_MARCA", 0).ToString().Trim());
                                dsDETA.SetValue("U_EXC_MODE", oEvent.Row - 1, oDataTable.GetValue("U_EXC_MODELO", 0).ToString().Trim());
                                dsDETA.SetValue("U_EXC_SEQP", oEvent.Row - 1, oDataTable.GetValue("SysNumber", 0).ToString());

                                ////FIRMWARE
                                //dsDETA.SetValue("U_EXC_FWAR", oEvent.Row - 1, oDataTable.GetValue("U_EXC_FIRMW", 0).ToString());
                                //FOTA
                                dsDETA.SetValue("U_EXC_FOTA", oEvent.Row - 1, oDataTable.GetValue("U_EXC_FOTA", 0).ToString());
                                //NRO SERIE LOTE
                                dsDETA.SetValue("U_EXC_NSLO", oEvent.Row - 1, oDataTable.GetValue("U_EXC_DISTNUM", 0).ToString());

                                oMatrix.LoadFromDataSource();

                                break;
                            case "Col_10": //IMEI CHIP
                                oMatrix.FlushToDataSource();
                                string IMEI = oDataTable.GetValue("U_EXC_IMEI", 0).ToString().Trim();
                                IMEI = string.IsNullOrEmpty(IMEI) ? oDataTable.GetValue("U_EXC_SIMCARD", 0).ToString().Trim() : IMEI;
                                dsDETA.SetValue("U_EXC_IMEI", oEvent.Row - 1, IMEI);
                                dsDETA.SetValue("U_EXC_OPER", oEvent.Row - 1, oDataTable.GetValue("U_EXC_OPERAD", 0).ToString().Trim());
                                dsDETA.SetValue("U_EXC_SCHI", oEvent.Row - 1, oDataTable.GetValue("SysNumber", 0).ToString());


                                //NUEVOS CAMPOS
                                dsDETA.SetValue("U_EXC_TIIP", oEvent.Row - 1, oDataTable.GetValue("U_EXC_TIPIP", 0).ToString());
                                dsDETA.SetValue("U_EXC_NRIP", oEvent.Row - 1, oDataTable.GetValue("U_EXC_IP", 0).ToString());
                                dsDETA.SetValue("U_EXC_PQDA", oEvent.Row - 1, oDataTable.GetValue("U_EXC_PAQDATOS", 0).ToString());
                                dsDETA.SetValue("U_EXC_LINE", oEvent.Row - 1, oDataTable.GetValue("U_EXC_LINEA", 0).ToString());
                                dsDETA.SetValue("U_EXC_LTEL", oEvent.Row - 1, oDataTable.GetValue("U_EXC_LINTEL", 0).ToString());
                                dsDETA.SetValue("U_EXC_DAPN", oEvent.Row - 1, oDataTable.GetValue("U_EXC_APN", 0).ToString());
                                dsDETA.SetValue("U_EXC_SIMC", oEvent.Row - 1, oDataTable.GetValue("U_EXC_SIMCARD", 0).ToString());
                                dsDETA.SetValue("U_EXC_PROT", oEvent.Row - 1, oDataTable.GetValue("U_EXC_PROTC", 0).ToString());

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

                        oMatrix.FlushToDataSource();
                        oMatrix.AutoResizeColumns();
                    }
                }


            }
            catch (Exception e)
            {
                result = false;
                StatusMessageError("whenChooseFromList > " + e.Message);
            }

            return result;
        }

        private void FiltrarIMEIChip(ItemEvent oEvent, ref bool result)
        {
            result = true;

            try
            {
                string itemCodeChip = ((EditText)oMatrix.GetCellSpecific("Col_7", oEvent.Row)).Value;
                if (string.IsNullOrEmpty(itemCodeChip))
                    throw new Exception("Debe elegir un Chip antes de registrar el IMEI Chip");

                SAPbouiCOM.Conditions oConditions = null;
                SAPbouiCOM.Condition oCondition = null;
                SAPbouiCOM.ChooseFromList oChooseFromList = null;
                oChooseFromList = mForm.ChooseFromLists.Item("cflICHI");
                oChooseFromList.SetConditions(null);

                oConditions = oChooseFromList.GetConditions();
                oCondition = oConditions.Add();
                oCondition.Alias = "ItemCode";
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCondition.CondVal = itemCodeChip;//SimCards
                oChooseFromList.SetConditions(oConditions);
            }
            catch (Exception ex)
            {
                result = false;
                Conexion.application.MessageBox(ex.Message);
            }
        }

        private void FiltrarIMEIEquipo(ItemEvent oEvent, ref bool result)
        {
            result = true;

            try
            {
                string itemcodeEquipo = ((EditText)oMatrix.GetCellSpecific("Col_2", oEvent.Row)).Value;
                if (string.IsNullOrEmpty(itemcodeEquipo))
                    throw new Exception("Debe elegir un equipo antes de registrar el IMEI Equipo");

                SAPbouiCOM.Conditions oConditions = null;
                SAPbouiCOM.Condition oCondition = null;
                SAPbouiCOM.ChooseFromList oChooseFromList = null;
                oChooseFromList = mForm.ChooseFromLists.Item("cflIEQP");
                oChooseFromList.SetConditions(null);

                oConditions = oChooseFromList.GetConditions();
                oCondition = oConditions.Add();
                oCondition.Alias = "ItemCode";
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCondition.CondVal = itemcodeEquipo;//SimCards
                oChooseFromList.SetConditions(oConditions);
            }
            catch (Exception ex)
            {
                result = false;
                Conexion.application.MessageBox(ex.Message);
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
                    //InstanciateCombo((ComboBox)oMatrix.GetCellSpecific("Col_4", 0), Queries.GetSerieItem(oRS.Fields.Item("Code").Value.ToString()), false);
                    RefreshChooseFromlist("cflIEQP", oRS.Fields.Item("Code").Value.ToString(), false);

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

        private string GetAlmacenOT(string DocEntry)
        {
            SAPbobsCOM.Recordset oRS = null;
            try
            {
                oRS = (SAPbobsCOM.Recordset)Conexion.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRS.DoQuery(Queries.GetAlmacenOT(DocEntry));
                if (oRS.RecordCount > 0)
                {
                    return oRS.Fields.Item("WareHouse").Value.ToString();
                }
                else return string.Empty;
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

        private void RefreshChooseFromlist(string cflCode, string ItemCode, bool isChip)
        {
            SAPbouiCOM.ChooseFromListCollection oChooseFromListCollection = mForm.ChooseFromLists;
            SAPbouiCOM.Conditions oConditions = null;
            SAPbouiCOM.Condition oCondition = null;
            SAPbouiCOM.ChooseFromList oChooseFromList = null;
            oChooseFromList = oChooseFromListCollection.Item(cflCode);
            oChooseFromList.SetConditions(oConditions);
            oConditions = oChooseFromList.GetConditions();
            oCondition = oConditions.Add();
            oCondition.Alias = "ItemCode";
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCondition.CondVal = ItemCode;//            
            oCondition.Relationship = BoConditionRelationship.cr_AND;
            oCondition = oConditions.Add();
            oCondition.Alias = "U_EXC_IMEI";
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
            oCondition.CondVal = "";//     

            if (isChip)
            {
                oCondition.Relationship = BoConditionRelationship.cr_AND;
                oCondition = oConditions.Add();
                oCondition.Alias = "U_EXC_OPERAD";
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                oCondition.CondVal = "";//     
            }
            else
            {
                oCondition.Relationship = BoConditionRelationship.cr_AND;
                oCondition = oConditions.Add();
                oCondition.Alias = "U_EXC_MARCA";
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                oCondition.CondVal = "";//     
                oCondition.Relationship = BoConditionRelationship.cr_AND;
                oCondition = oConditions.Add();
                oCondition.Alias = "U_EXC_MODELO";
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                oCondition.CondVal = "";//    
            }
            oChooseFromList.SetConditions(oConditions);
        }

        private void SetDetalleSerItem(string ItemCode, string Serie, int row, bool isChip = true)
        {
            SAPbobsCOM.Recordset oRS = null;
            try
            {
                oRS = (SAPbobsCOM.Recordset)Conexion.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRS.DoQuery(Queries.GetDetalleItemEnsamble(ItemCode, Serie));

                if (oRS.RecordCount > 0)
                {
                    if (isChip)
                    {
                        dsDETA.SetValue("U_EXC_IMEI", row, oRS.Fields.Item("U_EXC_IMEI").Value.ToString());
                        dsDETA.SetValue("U_EXC_OPER", row, oRS.Fields.Item("U_EXC_OPERAD").Value.ToString());
                    }
                    else
                    {
                        dsDETA.SetValue("U_EXC_MARC", row, oRS.Fields.Item("U_EXC_MARCA").Value.ToString());
                        dsDETA.SetValue("U_EXC_MODE", row, oRS.Fields.Item("U_EXC_MODELO").Value.ToString());
                        dsDETA.SetValue("U_EXC_IMEIEQP", row, oRS.Fields.Item("U_EXC_IMEI").Value.ToString());
                    }
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
                if (dsHEAD.GetValue("U_EXC_ESTA", 0) == "O")
                {
                    oMatrix.FlushToDataSource();
                    if (oMatrix.RowCount == 0) dsDETA.Clear();
                    dsDETA.InsertRecord(dsDETA.Size);
                    dsDETA.SetValue("LineId", dsDETA.Size - 1, dsDETA.Size.ToString());
                    oMatrix.LoadFromDataSource();
                    oMatrix.AutoResizeColumns();
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("#").Cells.Item(dsDETA.Size).Specific).Value = dsDETA.Size.ToString();
                }
                else
                {
                    StatusMessageWarning("Solo puede agregar filas en estado abierto.");
                }
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
                    SLDocument Excel = null;

                    if (Conexion.application.ClientType == BoClientType.ct_Desktop)
                    {
                        openFileDialog = new FolderFileDialog();
                        Archivo = openFileDialog.FindFile();
                        Excel = new SLDocument(Archivo);
                    }
                    else if (Conexion.application.ClientType == BoClientType.ct_Browser)
                    {
                        Archivo = ((SAPbouiCOM.EditText)mForm.Items.Item("RUTA").Specific).Value;

                        if (string.IsNullOrEmpty(Archivo)) throw new Exception("Debe ingresar la ruta de un archivo excel");
                        else
                        {
                            bool esRuta = System.IO.Path.IsPathRooted(Archivo) && Archivo.IndexOfAny(System.IO.Path.GetInvalidPathChars()) == -1;
                            if (!esRuta) throw new Exception("La ruta del archivo excel no es válida");
                            else
                            {
                                if (!File.Exists(Archivo)) throw new Exception("El archivo excel no existe en la ruta indicada");
                                else
                                {
                                    string extension = System.IO.Path.GetExtension(Archivo).ToLower();
                                    if(!string.Equals(extension, ".xlsx", StringComparison.OrdinalIgnoreCase)) 
                                        throw new Exception("El archivo de la ruta debe ser un excel con extensión .xlsx");
                                }
                            }
                        }

                        Excel = new SLDocument(Archivo);
                    }

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
                        //var Excel = new SLDocument(Archivo);
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
                            ((EditText)oMatrix.Columns.Item("Col_17").Cells.Item(oMatrix.VisualRowCount).Specific).Value = row["IMEI"].ToString();
                            if (string.IsNullOrEmpty(row["IMEI"].ToString())) ((EditText)oMatrix.Columns.Item("Col_14").Cells.Item(oMatrix.VisualRowCount).Specific).Value = "";
                            ((EditText)oMatrix.Columns.Item("Col_7").Cells.Item(oMatrix.VisualRowCount).Specific).Value = row["CHIP"].ToString();
                            ((EditText)oMatrix.Columns.Item("Col_10").Cells.Item(oMatrix.VisualRowCount).Specific).Value = row["IMEI CHIP (SimCard)"].ToString();
                            ((EditText)oMatrix.Columns.Item("Col_12").Cells.Item(oMatrix.VisualRowCount).Specific).Value = row["VERSION (Firmware)"].ToString();
                            ((EditText)oMatrix.Columns.Item("Col_13").Cells.Item(oMatrix.VisualRowCount).Specific).Value = row["Archivo de Configuracion"].ToString();
                            ((EditText)oMatrix.Columns.Item("Col_18").Cells.Item(oMatrix.VisualRowCount).Specific).Value = row["Servicio"].ToString();
                            ((EditText)oMatrix.Columns.Item("Col_14").Cells.Item(oMatrix.VisualRowCount).Specific).Value = row["Producido Por"].ToString();
                            ((EditText)oMatrix.Columns.Item("Col_30").Cells.Item(oMatrix.VisualRowCount).Specific).Value = row["Accesorio"].ToString();
                            ((EditText)oMatrix.Columns.Item("Col_32").Cells.Item(oMatrix.VisualRowCount).Specific).Value = row["Accesorio 2"].ToString();
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