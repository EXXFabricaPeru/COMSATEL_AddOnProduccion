using AddonProduccionEnsDes.commons;
using AddonProduccionEnsDes.conexion;
using AddonProduccionEnsDes.data_schema;
using System;
using System.Collections.Generic;

namespace AddonProduccionEnsDes.view
{
    public class frmConfig : FormCommon, IForm
    {
        #region variables
        private SAPbouiCOM.Form mForm;
        private SAPbouiCOM.DBDataSource dsConfig;
        private SAPbouiCOM.DBDataSource dsTasas;
        private SAPbouiCOM.Matrix oMatrix;
        private SAPbouiCOM.ComboBox cboCodTax, cboCodExp;

        private const string EDT_CUENTA = "edtAct";
        private const string EDT_ITEM = "edtItem";
        private const string EDT_ENTRY = "edtEnt";
        private const string BTN_ADD = "btnAdd";
        private const string MTX_FACTOR = "mtxFact";
        string code = "1";

        //Right Click
        private string ItemUIDRightClick;
        private int RowItemRightClick;
        #endregion

        public frmConfig(Dictionary<string, IForm> dictionary)
        {
            mForm = CreateForm(Conexion.company, Conexion.application, Properties.Resources.frmCfgProIGV, FormName.CONFIG_PI);
            if (mForm != null)
            {
                dictionary.Add(getFormUID(), this);
                Initializer();
                mForm.Visible = true;
            }
            else
                StatusMessageError("Construct: No se pudo crear el formulario " + FormName.CONFIG_PI);
        }

        #region _EVENTOS_ITEMEVENT

        //Principal

        private void Initializer()
        {
            try
            {
                mForm.Freeze(true);
                dsConfig = mForm.DataSources.DBDataSources.Item("@" + SCConfig.TABLE_CABE);

                cboCodTax = mForm.Items.Item("cboTax").Specific;
                cboCodExp = mForm.Items.Item("cboAdd").Specific;
                LoadDefaults();
                //AddRow();
                //oMatrix.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                StatusMessageError("cargarOpcionesPorDefecto > " + ex.Message);
            }
            finally { mForm.Freeze(false); }
        }
        private void LoadDefaults()
        {
            LoadCombo();
            if (!string.IsNullOrEmpty(GetConfig("Code"))) mForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
            mForm.Items.Item(EDT_ENTRY).Enabled = true;
            mForm.Items.Item(EDT_ENTRY).Specific.Value = code;
            if (!string.IsNullOrEmpty(GetConfig("Code"))) mForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            mForm.Items.Item(EDT_ENTRY).Visible = false;
        }
        private void LoadCombo()
        {
            try
            {
                cboCodTax.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                cboCodExp.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                InstanciateCombo(cboCodTax, Queries.GetComboTax());
                //InstanciateCombo(cboCodExp, Queries.GetComboExp());
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
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                        result = WhenDataAdd(itemEvent);
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
        private void WhenLostFocus(SAPbouiCOM.ItemEvent oEvent)
        {
            switch (oEvent.ItemUID)
            {
                default:
                    break;
            }
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
            switch (oEvent.ItemUID)
            {
                case BTN_ADD:
                    if (!oEvent.BeforeAction)
                        AddRow();
                    break;
                case "1":
                    //dsConfig.SetValue("Code", 0, code);
                    break;
                default:
                    break;
            }
            return res;
        }

        private bool WhenDataAdd(SAPbouiCOM.ItemEvent oEvent)
        {
            bool res = true;
            switch (oEvent.ItemUID)
            {
                default:
                    break;
            }
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
                        case EDT_CUENTA:
                            dsConfig.SetValue("U_FormatCode", 0, oDataTable.GetValue("AcctCode", 0).Trim());
                            //dsConfig.SetValue("U_MSST_NEM", 0, oDataTable.GetValue("AcctName", 0).Trim());
                            break;
                        case EDT_ITEM:
                            dsConfig.SetValue("U_ItemCode", 0, oDataTable.GetValue("ItemCode", 0).Trim());
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
        public bool HandleFormDataEvents(SAPbouiCOM.BusinessObjectInfo oBusinessObjectInfo)
        {
            switch (oBusinessObjectInfo.EventType)
            {
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
                    switch (menuEvent.MenuUID)
                    {
                        case Constants.Menu_EliminarLinea:
                            DeleteRow(RowItemRightClick, ItemUIDRightClick);
                            break;
                        case Constants.Menu_AgregarLinea:
                            AddRow();
                            break;
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
                        oCreationPackage = Conexion.application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                        oMenuItem = Conexion.application.Menus.Item(Constants.Menu_Context);
                        oMenus = oMenuItem.SubMenus;

                        ItemUIDRightClick = menuInfo.ItemUID;
                        RowItemRightClick = menuInfo.Row;

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
                if (ItemUID.Equals(MTX_FACTOR))
                {
                    oMatrix = mForm.Items.Item(ItemUID).Specific;
                    oMatrix.FlushToDataSource();


                    dsTasas.RemoveRecord(row - 1);

                    if (mForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        mForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    oMatrix.LoadFromDataSource();

                }
            }
            catch (Exception ex)
            {
                StatusMessageError("DeleteRow() > " + ex.Message);
            }
        }
        private void AddRow()
        {
            try
            {
                mForm.Freeze(true);
                oMatrix = mForm.Items.Item(MTX_FACTOR).Specific;
                oMatrix.FlushToDataSource();
                if (oMatrix.RowCount == 0) { dsTasas.Clear(); }
                dsTasas.InsertRecord(dsTasas.Size);
                dsTasas.SetValue("U_Fecha", dsTasas.Size - 1, GetFecha());
                dsTasas.SetValue("U_Factor", dsTasas.Size - 1, string.Empty);
                dsTasas.SetValue("U_FactReal", dsTasas.Size - 1, string.Empty);
                oMatrix.LoadFromDataSource();
                oMatrix.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                StatusMessageError("AddRow() > " + ex.Message);
            }

            finally { mForm.Freeze(false); }
        }

        private string GetFecha()
        {
            string fecha = DateTime.Now.ToString(Constants.PERDATE_FORMAT);
            return fecha;

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