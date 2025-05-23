
using AddonProduccionEnsDes.commons;
using AddonProduccionEnsDes.conexion;
using System;
using System.Globalization;

namespace AddonProduccionEnsDes.view
{
    public class frmOCTR : FormCommon, IForm
    {
        #region variables
        private SAPbouiCOM.Form mForm;
        private SAPbouiCOM.Matrix oMatrix;
        private SAPbouiCOM.Grid oGrid;
        private SAPbouiCOM.DataTable oDTRec, oDTRecDet;
        private const string OBJECT_CODE = "OPDN";

        private const string GRD_ITM = "38", GRD_SERV = "39";//ID Grid
        private const string BTN_OK = "1"; //ID BOTONES
        private const string BTN_FORMUL = "btnForm", TXT_CONENTRY = "10";//FormButtom
        private const string TAB_PERIOD = "540000151", MTX_MODELOTXT = "540000152", MTX_PERIODTXT = "540000153", MTX_MODELOSYS = "540000154", MTX_PERIODSYS = "540000155";//ExitingItems
        private const string MTX_REC = "mtxRec", MTX_RECDET = "mtxRecDet", DT_REC = "DTRecurrente", DT_RECDET = "DTRecurrenteDet";

        #endregion

        public frmOCTR() { }


        #region _EVENTOS_ITEMEVENT

        //Principal
        public bool HandleItemEvents(SAPbouiCOM.ItemEvent itemEvent)
        {
            var result = true;
            try
            {
                switch (itemEvent.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                        result = WhenFormLoad(itemEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
                        result = WhenFormResize(itemEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        result = WhenItemPressed(itemEvent);
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
        private bool WhenLostFocus(SAPbouiCOM.ItemEvent oEvent)
        {
            bool res = true;

            switch (oEvent.ItemUID)
            {
                default:
                    break;
            }
            return res;
        }


        private bool WhenFormLoad(SAPbouiCOM.ItemEvent oEvent)
        {
            if (!oEvent.BeforeAction) return true;
            try
            {
                Conexion.formOpen.Add(oEvent.FormUID, (commons.IForm)this);
                mForm = Conexion.application.Forms.Item(oEvent.FormUID);
            }
            catch (Exception)
            {
            }
            return true;
        }
        private bool WhenFormResize(SAPbouiCOM.ItemEvent oEvent)
        {
            if (!oEvent.BeforeAction)
            {
                mForm = Conexion.application.Forms.Item(oEvent.FormUID);
                RefrescarForm();
            }

            return true;
        }
        private bool WhenItemPressed(SAPbouiCOM.ItemEvent oEvent)
        {
            bool res = true;
            switch (oEvent.ItemUID)
            {
                case TAB_PERIOD:


                    if (!oEvent.BeforeAction && oEvent.ActionSuccess)
                    {
                        res = ActualizarLayout(oEvent);
                        if (res) res = ActualizarRecurrente();
                        if (res) res = ActualizarRecurrentesDet();
                    }
                    break;
                case MTX_REC:
                    if (!oEvent.BeforeAction)
                    {
                        mForm = Conexion.application.Forms.Item(oEvent.FormUID);
                        mForm.Freeze(true);
                        try
                        {
                            res = ActualizarRecurrentesDet(false, oEvent.Row);
                            FormatGrids(false);
                        }
                        catch (Exception ex)
                        {
                        }
                        mForm.Freeze(false);


                    }
                    break;
                default:
                    break;
            }
            return res;
        }
        private void RefrescarForm()
        {
            mForm.Freeze(true);
            try
            {
                var matrix = mForm.Items.Item(MTX_REC);
                matrix.Left = mForm.Items.Item(MTX_MODELOSYS).Left;
                matrix.Width = mForm.Items.Item(MTX_MODELOSYS).Width + 140;
                matrix = mForm.Items.Item(MTX_RECDET);
                matrix.Left = mForm.Items.Item(MTX_PERIODSYS).Left + 140;
                matrix.Width = mForm.Items.Item(MTX_PERIODSYS).Width - 140;
                mForm.Items.Item(MTX_PERIODTXT).Left = matrix.Left;
            }
            catch (Exception)
            {
            }
            mForm.Freeze(false);
        }

        private bool ActualizarRecurrente(bool isFormat = false)
        {

            oDTRec = mForm.DataSources.DataTables.Item(DT_REC);
            oDTRec.ExecuteQuery(Queries.GetRecurrentes(isFormat ? "-1" : mForm.Items.Item(TXT_CONENTRY).Specific.Value));
            if (oDTRec.Rows.Count > 0)
                FormatGrids();
            return true;
        }

        private void FormatGrids(bool isRec = true)
        {

            SAPbouiCOM.Item oGeneric;
            SAPbouiCOM.Grid oGrid;
            SAPbouiCOM.Column oColumn;
            SAPbouiCOM.EditTextColumn col;
            if (isRec)
            {
                oGeneric = mForm.Items.Item(MTX_REC);
                oGeneric.Enabled = false;
                oGrid = oGeneric.Specific;
                col = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("DocEntry");
                col.Visible = false;
                col = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("Base");
                col.LinkedObjectType = "17";
            }
            else
            {
                oGeneric = mForm.Items.Item(MTX_RECDET);
                oGeneric.Enabled = false;
                oGrid = oGeneric.Specific;
                oGrid.AutoResizeColumns();
                col = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("Instancia");
                col.LinkedObjectType = "17";
            }

        }
        private bool ActualizarRecurrentesDet(bool isFormat = false, int row = -1)
        {
            oDTRecDet = mForm.DataSources.DataTables.Item("DTRecurrenteDet");
            oDTRecDet.ExecuteQuery(Queries.GetDetalleRecurrente(isFormat ? "-1" : SelectedMtxRow(mForm.Items.Item(MTX_REC).Specific, row)));
            return true;
        }

        private string SelectedMtxRow(SAPbouiCOM.Grid oGrid, int row)
        {
            int entry = -1;
            try
            {
                if (oGrid.Rows.SelectedRows.Count > 0)
                {
                    entry = oGrid.DataTable.GetValue("DocEntry", row);
                }
            }
            catch (Exception)
            {

            }

            return entry.ToString();
        }

        private bool ActualizarLayout(SAPbouiCOM.ItemEvent oEvent)
        {
            bool res = true;
            mForm = Conexion.application.Forms.Item(oEvent.FormUID);
            mForm.Freeze(true);
            if (mForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) res = false;
            RemoveUIAux();
            AddUIAux();

            mForm.Freeze(false);

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

        public bool HandleFormDataEvents(SAPbouiCOM.BusinessObjectInfo oBusinessObjectInfo)
        {
            switch (oBusinessObjectInfo.EventType)
            {
                case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                    WhenFormDataLoad(oBusinessObjectInfo);
                    break;
                default:
                    break;
            }
            return true;
        }

        private void WhenFormDataLoad(SAPbouiCOM.BusinessObjectInfo oBusinessObjectInfo)
        {
            mForm = Conexion.application.Forms.Item(oBusinessObjectInfo.FormUID);
            ActualizarRecurrente();
            ActualizarRecurrentesDet();
        }

        public bool HandleMenuDataEvents(SAPbouiCOM.MenuEvent menuEvent)
        {
            return true;
        }

        #region _EVENTS_RIGHTCLICK
        public bool HandleRightClickEvent(SAPbouiCOM.ContextMenuInfo menuInfo)
        {
            var result = true;
            return result;
        }
        #endregion

        #region _METODOS_PROPIOS

        private void RemoveUIAux()
        {
            SAPbouiCOM.Items oItems = mForm.Items;
            oItems.Item(MTX_MODELOTXT).Visible = false;
            oItems.Item(MTX_MODELOSYS).Visible = false;
            oItems.Item(MTX_PERIODSYS).Visible = false;

        }
        private void AddUIAux()
        {
            SAPbouiCOM.Items oItems = mForm.Items;
            SAPbouiCOM.Item oGeneric;
            SAPbouiCOM.Grid oGrid;
            try
            {
                oGeneric = oItems.Add(MTX_REC, SAPbouiCOM.BoFormItemTypes.it_GRID);
                oGeneric.Top = oItems.Item(MTX_MODELOSYS).Top;
                oGeneric.Left = oItems.Item(MTX_MODELOSYS).Left;
                oGeneric.Width = oItems.Item(MTX_MODELOSYS).Width + 140;
                oGeneric.Height = oItems.Item(MTX_MODELOSYS).Height;
                oGeneric.FromPane = 7;
                oGeneric.ToPane = 7;
                oGeneric.Enabled = true;
                oGeneric.Visible = true;
                oGrid = oGeneric.Specific;
                oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;
                oDTRec = mForm.DataSources.DataTables.Add(DT_REC);
                oGrid.DataTable = ((SAPbouiCOM.DataTable)(oDTRec));
                ActualizarRecurrente(true);

                oGeneric = oItems.Add(MTX_RECDET, SAPbouiCOM.BoFormItemTypes.it_GRID);
                oGeneric.Top = oItems.Item(MTX_PERIODSYS).Top;
                oGeneric.Left = oItems.Item(MTX_PERIODSYS).Left + 140;
                oGeneric.Width = oItems.Item(MTX_PERIODSYS).Width - 140;
                oGeneric.Height = oItems.Item(MTX_PERIODSYS).Height;
                oGeneric.FromPane = 7;
                oGeneric.ToPane = 7;
                oGeneric.Enabled = true;
                oGeneric.Visible = true;
                oGrid = oGeneric.Specific;
                oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_None;
                oDTRecDet = mForm.DataSources.DataTables.Add(DT_RECDET);
                oGrid.DataTable = ((SAPbouiCOM.DataTable)(oDTRecDet));
                ActualizarRecurrentesDet(true);

            }
            catch (Exception ex)
            {
            }
            finally
            {
                LiberarObjetoGenerico(oItems);
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