
using AddonProduccionEnsDes.commons;
using AddonProduccionEnsDes.conexion;
using System;
using System.Globalization;

namespace AddonProduccionEnsDes.view
{
    public class frmOCSL : FormCommon, IForm
    {
        #region variables
        private SAPbouiCOM.Form mForm;
        private SAPbouiCOM.Matrix oMatrix;
        private SAPbouiCOM.Grid oGrid;
        private SAPbouiCOM.DataTable oDTDet, oDTRecDet;
        private const string OBJECT_CODE = "OPDN";

        private const string GRD_ITM = "38", GRD_SERV = "39";//ID Grid
        private const string BTN_OK = "1"; //ID BOTONES
        private const string BTN_FORMUL = "btnForm", TXT_ENTRY = "12";//FormButtom
        private const string TAB_PERIOD = "540000151", MTX_MODELOTXT = "540000152", MTX_PERIODTXT = "540000153", GRD_ANEXOS = "234000001", GRD_DET = "GRD_DET", MTX_PERIODSYS = "540000155", TAB_ANX = "234000005", TAB_DET = "tabDet";//ExitingItems
        private const string MTX_REC = "mtxRec", MTX_RECDET = "mtxRecDet", DT_DET = "DTLLDetalle", DT_RECDET = "DTRecurrenteDet";

        #endregion

        public frmOCSL() { }


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
            if (oEvent.BeforeAction) {
                mForm = Conexion.application.Forms.Item(oEvent.FormUID);
                ActualizarLayout(oEvent);
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
                case TAB_DET:
                    mForm = Conexion.application.Forms.Item(oEvent.FormUID);
                    mForm.PaneLevel = 11;
                    if (res) res = ActualizarDetalleLlamada();

                    break;
                case TAB_PERIOD:
                    if (!oEvent.BeforeAction && oEvent.ActionSuccess)
                    {
                        //res = ActualizarLayout(oEvent);
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
                var matrix = mForm.Items.Item(TAB_DET);
                matrix.Left = mForm.Items.Item(TAB_ANX).Left;
                matrix.Top = mForm.Items.Item(TAB_ANX).Top;
            }
            catch (Exception)
            {
            }
            mForm.Freeze(false);
        }

        private bool ActualizarDetalleLlamada(bool isFormat = false)
        {

            oDTDet = mForm.DataSources.DataTables.Item(DT_DET);
            oDTDet.ExecuteQuery(Queries.GetDetalleLlamada(isFormat ? "-1" : mForm.Items.Item(TXT_ENTRY).Specific.Value));
            if (oDTDet.Rows.Count > 0)
                FormatGrids();
            return true;
        }

        private void FormatGrids()
        {

            SAPbouiCOM.Item oGeneric;
            SAPbouiCOM.Grid oGrid;
            SAPbouiCOM.Column oColumn;
            SAPbouiCOM.EditTextColumn col;

                //oGeneric = mForm.Items.Item(MTX_REC);
                //oGeneric.Enabled = false;
                //oGrid = oGeneric.Specific;
                //col = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("DocEntry");
                //col.Visible = false;
                //col = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("Base");
                //col.LinkedObjectType = "17";

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
            mForm.Freeze(true);
            if (mForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) res = false;
            //RemoveUIAux();
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
            ActualizarDetalleLlamada();
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
            oItems.Item(GRD_ANEXOS).Visible = false;
            oItems.Item(MTX_PERIODSYS).Visible = false;

        }
        private void AddUIAux()
        {
            SAPbouiCOM.Items oItems = mForm.Items;
            SAPbouiCOM.Item oItm;
            SAPbouiCOM.Folder oFolder;
            SAPbouiCOM.Grid oGrid;
            try
            {
                oItm = oItems.Add(GRD_DET, SAPbouiCOM.BoFormItemTypes.it_GRID);
                oItm.Top = oItems.Item(GRD_ANEXOS).Top;
                oItm.Left = oItems.Item(GRD_ANEXOS).Left;
                oItm.Width = oItems.Item(GRD_ANEXOS).Width;
                oItm.Height = oItems.Item(GRD_ANEXOS).Height;
                oItm.FromPane = 11;
                oItm.ToPane = 11;
                oItm.Enabled = true;

                oGrid = oItm.Specific;
                oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_None;

                oDTDet = mForm.DataSources.DataTables.Add(DT_DET);
                oGrid.DataTable = ((SAPbouiCOM.DataTable)(oDTDet));
                ActualizarDetalleLlamada(true);
                oItm = oItems.Add(TAB_DET, SAPbouiCOM.BoFormItemTypes.it_FOLDER);
                oItm.Top = oItems.Item(TAB_ANX).Top;
                oItm.Height = oItems.Item(TAB_ANX).Height;
                oItm.Width = oItems.Item(TAB_ANX).Width;
                oItm.Left = oItems.Item(TAB_ANX).Left+ oItems.Item(TAB_ANX).Width;
                oFolder = oItm.Specific;
                oFolder.Caption = "Detalle";
                oFolder.GroupWith(TAB_ANX);
                oItm.Visible = true;
                mForm.PaneLevel = 1;
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