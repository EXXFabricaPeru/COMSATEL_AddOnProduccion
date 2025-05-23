using AddonProduccionEnsDes.commons;
using AddonProduccionEnsDes.conexion;
using System;
using System.Globalization;

namespace AddonProduccionEnsDes.view
{
    public class frmOITT : FormCommon, IForm
    {
        #region variables
        private SAPbouiCOM.Form mForm;
        private SAPbouiCOM.Matrix oMatrix;

        private const string OBJECT_CODE = "OPDN";

        private const string GRD_ITM = "38", GRD_SERV = "39";//ID Grid
        private const string BTN_OK = "1"; //ID BOTONES
        private const string BTN_FORMUL = "btnForm", BASE_POS = "540000039";//FormButtom
        private const string MTX_LISTA = "3", EDT_ITEMCODE = "4", EDT_WHS = "16", EDT_QTY = "17", EDT_DOCDATE = "10", EDT_NUMATCARD = "14", EDT_TYPE = "3", EDT_CURR = "63";//FormLogic
        private const string COL_TAX = "160", COL_TAXVAL = "19", COL_WHS = "3", COL_QTY = "11", COL_PRICETAX = "20", COL_DSC = "15", COL_GASTCOD = "111", COL_GASTVAL = "112", COL_GASTVALFC = "113";//ITMS
        private const string COL_TAX_SERV = "95", COL_TAXVAL_SERV = "10", COL_PRICE_SERV = "14";//SERV
        private const string COL_TIPO = "1880000002", COL_CODE = "1", COL_DESC = "44", COL_CANT = "2", COL_PRICE = "4", COL_CANTADIC = "1880000005", COL_POSICION = "U_EXP_POSICION", COL_PREDECE = "U_EXP_PREDECE";//LISTA MAT
        private const int COL_ITM = 5, COL_STOCK = 2;

        #endregion

        public frmOITT() { }


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
            AddUIAux(Conexion.application.Forms.Item(oEvent.FormUID));
            return true;
        }

        private bool WhenItemPressed(SAPbouiCOM.ItemEvent oEvent)
        {
            bool res = true;
            switch (oEvent.ItemUID)
            {
                case BTN_FORMUL:
                    if (!oEvent.BeforeAction)
                        res = OperateFormulacion(oEvent);
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

        private void AddUIAux(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Items oItems = oForm.Items;
            SAPbouiCOM.Item oGeneric;
            SAPbouiCOM.Button oBtn;
            try
            {
                oGeneric = oItems.Add(BTN_FORMUL, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oGeneric.Top = oItems.Item(BASE_POS).Top + oItems.Item(BASE_POS).Height + 10;
                oGeneric.Left = oItems.Item(BASE_POS).Left;
                oGeneric.Width = oItems.Item(BASE_POS).Width;
                oGeneric.Enabled = true;
                oGeneric.Visible = true;
                oBtn = oGeneric.Specific;
                oBtn.Caption = "Formulación";
            }
            catch (Exception ex)
            {
                StatusMessageError("AddUIExcel() > " + ex.Message);
            }
            finally
            {
                LiberarObjetoGenerico(oItems);
            }
        }

        private bool OperateFormulacion(SAPbouiCOM.ItemEvent oEvent)
        {
            bool res = true;
            mForm = Conexion.application.Forms.Item(oEvent.FormUID);
            string itemCode = mForm.Items.Item(EDT_ITEMCODE).Specific.value;
            oMatrix = mForm.Items.Item(MTX_LISTA).Specific;

            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)Conexion.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                if (mForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    StatusMessageWarning("Solo se formula en modo creación");
                    return false;
                }
                if (string.IsNullOrEmpty(itemCode))
                {
                    StatusMessageWarning("No se tiene ItemCode");
                    return false;
                }
                oRS.DoQuery(Queries.GetFormulacion(itemCode));

                if (oRS.RecordCount == 0)
                {
                    StatusMessageWarning("No hay formulación asignada");
                    return false;
                }
                StatusMessageSuccess("Calculando formulas");
                oRS.DoQuery(Queries.GetFormulacionDetalle(itemCode));
                if (oRS.RecordCount > 0)
                {
                    mForm.Freeze(true);

                    int position = 1;
                    mForm.Items.Item(EDT_QTY).Specific.value = oRS.Fields.Item("KILO").Value;
                    mForm.Items.Item(EDT_WHS).Specific.Select(oRS.Fields.Item("WHS").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                    while (!oRS.EoF)
                    {

                        oMatrix.Columns.Item(COL_TIPO).Cells.Item(position).Specific.Select(oRS.Fields.Item("TIPO").Value, SAPbouiCOM.BoSearchKey.psk_Index);
                        oMatrix.Columns.Item(COL_CODE).Cells.Item(position).Specific.Value = oRS.Fields.Item("ITEMCODE").Value;
                        if (oRS.Fields.Item("TIPO").Value == 3)
                        {
                            oMatrix.Columns.Item(COL_DESC).Cells.Item(position).Specific.Value = oRS.Fields.Item("DESC").Value;
                        }
                        else
                        {
                            oMatrix.Columns.Item(COL_CANT).Cells.Item(position).Specific.Value = oRS.Fields.Item("CANT").Value == 0 ? -1 : oRS.Fields.Item("CANT").Value;
                            oMatrix.Columns.Item(COL_CANTADIC).Cells.Item(position).Specific.Value = oRS.Fields.Item("CADI").Value;
                            if (oRS.Fields.Item("PRICE").Value != 0) oMatrix.Columns.Item(COL_PRICE).Cells.Item(position).Specific.Value = oRS.Fields.Item("PRICE").Value;
                            oMatrix.Columns.Item(COL_WHS).Cells.Item(position).Specific.Value = oRS.Fields.Item("WHS").Value;
                            oMatrix.Columns.Item(COL_POSICION).Cells.Item(position).Specific.Value = oRS.Fields.Item("RUTA").Value;
                            if (oRS.Fields.Item("PRED").Value != "0") oMatrix.Columns.Item(COL_PREDECE).Cells.Item(position).Specific.Value = oRS.Fields.Item("PRED").Value;
                        }

                        position++;
                        oRS.MoveNext();
                    }
                    mForm.Freeze(false);
                }

            }
            catch (Exception EX)
            {

                throw;
            }
            finally
            {
                mForm.Freeze(false);

            }

            return res;

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