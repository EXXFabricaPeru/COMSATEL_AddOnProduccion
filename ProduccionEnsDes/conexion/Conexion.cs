using AddonProduccionEnsDes.commons;
using AddonProduccionEnsDes.data_schema;
using AddonProduccionEnsDes.view;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Xml;

namespace AddonProduccionEnsDes.conexion
{
    public class Conexion
    {
        public static SAPbobsCOM.Company company;
        public static SAPbouiCOM.Application application;
        public static readonly Dictionary<string, IForm> formOpen;
        static Conexion()
        {
            formOpen = new Dictionary<string, IForm>();
        }
        public Conexion()
        {
            try
            {
                application = instanciarAplicacion();
                company = InstanciarCompania();
                InicializarFiltros();
                DataStructure sd = new DataStructure();
                application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(Application_AppEvent);
                application.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(Application_MenuEvent);
                application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(Application_ItemEvent);
                application.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(Application_FormDataEvent);
                application.RightClickEvent += new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(Application_RightClickEvent);

                //LAYOUT
                application.LayoutKeyEvent += new SAPbouiCOM._IApplicationEvents_LayoutKeyEventEventHandler(Application_LayoutKeyEvent);
                CrearMenu();
            }
            catch (Exception e)
            {
                application.MessageBox("Conexion: " + e.Message);
            }
        }

        private void Application_LayoutKeyEvent(ref SAPbouiCOM.LayoutKeyInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            //NO QUISE TOCAR TU INTERFAZ(IForm) PARA QUE ACEPTARA ESTE EVENTO, ASI QUE ESTE EVENTO AISLADO LO MANEJO SOLO ACA (DAVID)
            if (eventInfo.BeforeAction && eventInfo.FormUID.Contains(FormName.FORMUL_RLM))
            {
                SAPbouiCOM.Form form = Conexion.application.Forms.ActiveForm;

                string codigo = form.DataSources.DBDataSources.Item("@EXP_OFRM").GetValue("Code", 0);
                if (!string.IsNullOrEmpty(codigo))
                    eventInfo.LayoutKey = codigo;
                else BubbleEvent = false;
            }
        }

        private SAPbouiCOM.Application instanciarAplicacion()
        {
            var guiApi = new SAPbouiCOM.SboGuiApi();
            guiApi.Connect(Environment.GetCommandLineArgs().GetValue(1).ToString());
            return guiApi.GetApplication();
        }
        private SAPbobsCOM.Company InstanciarCompania()
        {
            try
            {
                //return application.Company.GetDICompany();
                SAPbobsCOM.Company sboCompany = new SAPbobsCOM.Company();
                string cookie = sboCompany.GetContextCookie();
                string conStr = application.Company.GetConnectionContext(cookie);

                if (sboCompany.Connected)
                    sboCompany.Disconnect();

                int ret = sboCompany.SetSboLoginContext(conStr);

                if (ret != 0)
                    throw new Exception("Login context failed");

                ret = sboCompany.Connect();

                return sboCompany;
            }
            catch (Exception e)
            {
                application.MessageBox(e.Message);
            }
            return null;
        }

        private void InicializarFiltros()
        {
            SAPbouiCOM.EventFilters filtros = new SAPbouiCOM.EventFilters();
            SAPbouiCOM.EventFilter filtroMenu = filtros.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK);
            SAPbouiCOM.EventFilter filtroItem = filtros.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED);


            filtroItem.AddEx(FormName.LLAMADA_SERVICIO);
            filtroItem.AddEx(FormName.CONTRATO);
            filtroItem.AddEx(FormName.ENSAMBLE);
            filtroItem.AddEx(FormName.DESENSAMBLE);
            filtroItem.AddEx("10010044"); //Lista de datos maestros de numero de serie

            SAPbouiCOM.EventFilter filtroFocus = filtros.Add(SAPbouiCOM.BoEventTypes.et_VALIDATE);
            filtroFocus.AddEx(FormName.ENSAMBLE);
            filtroFocus.AddEx(FormName.DESENSAMBLE);

            SAPbouiCOM.EventFilter filtroCFL = filtros.Add(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST);
            filtroCFL.AddEx(FormName.ENSAMBLE);
            filtroCFL.AddEx(FormName.DESENSAMBLE);

            //SAPbouiCOM.EventFilter filterMatrixLink = filtros.Add(SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED);
            //filterMatrixLink.AddEx(FormName.FORMUL_RLM);

            SAPbouiCOM.EventFilter filterCombo = filtros.Add(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT);
            filterCombo.AddEx(FormName.ENSAMBLE);
            filterCombo.AddEx(FormName.DESENSAMBLE);

            //SAPbouiCOM.EventFilter filterLostFocus = filtros.Add(SAPbouiCOM.BoEventTypes.et_LOST_FOCUS);
            //filterLostFocus.AddEx(FormName.CONTRATO);

            //SAPbouiCOM.EventFilter filterFormLoad = filtros.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD);
            //filterFormLoad.AddEx(FormName.ENSAMBLE);
            //filterFormLoad.AddEx(FormName.DESENSAMBLE);

            ////filterFormLoad.AddEx("0");
            //SAPbouiCOM.EventFilter filterAddData = filtros.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD);
            //filterAddData.AddEx(FormName.CONTRATO);
            SAPbouiCOM.EventFilter filterLoadData = filtros.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD);
            filterLoadData.AddEx(FormName.ENSAMBLE);
            filterLoadData.AddEx(FormName.DESENSAMBLE);

            //SAPbouiCOM.EventFilter filterFormResize = filtros.Add(SAPbouiCOM.BoEventTypes.et_FORM_RESIZE);
            //filterFormResize.AddEx(FormName.CONTRATO);
            //filterFormResize.AddEx(FormName.LLAMADA_SERVICIO);

            //SAPbouiCOM.EventFilter filterClose = filtros.Add(SAPbouiCOM.BoEventTypes.et_FORM_CLOSE);
            //filterClose.AddEx(FormName.DATOS_DEL_VIAJE);
            SAPbouiCOM.EventFilter filterRightClick = filtros.Add(SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK);
            filterRightClick.AddEx(FormName.ENSAMBLE);
            filterRightClick.AddEx(FormName.DESENSAMBLE);

            

            application.SetFilter(filtros);
        }

        //Eventos de aplicación
        void Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            try
            {
                BubbleEvent = true;

                if (formOpen.ContainsKey(FormUID))
                {
                    BubbleEvent = formOpen[FormUID].HandleItemEvents(pVal);
                }


                if (pVal.FormTypeEx == "10010044")
                {
                    if (pVal.BeforeAction && pVal.ItemUID == "5")
                    {
                        BubbleEvent = false;
                        Conexion.application.StatusBar.SetText(Constants.PREFIX_MSG_ADDON + "Operación no disponible", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                }
                //switch (pVal.FormTypeEx)
                //{
                //    case FormName.CONTRATO:
                //        BubbleEvent = new frmOCTR().HandleItemEvents(pVal);
                //        break;
                //    case FormName.LLAMADA_SERVICIO:
                //        BubbleEvent = new frmOCSL().HandleItemEvents(pVal);
                //        break;
                //    default:
                //        break;
                //}
            }
            catch (Exception)
            {
                BubbleEvent = true;
            }
        }

        void Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    company.Disconnect();
                    Environment.Exit(0);
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    company.Disconnect();
                    Environment.Exit(0);
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    company.Disconnect();
                    Environment.Exit(0);
                    break;
            }
        }

        void Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            var result = true;
            if (pVal.BeforeAction)
            {
                try
                {
                    switch (pVal.MenuUID)
                    {
                        case FormName.ENSAMBLE:
                            frmEnsamble ensamble = new frmEnsamble(formOpen);
                            break;
                        case FormName.DESENSAMBLE:
                            frmDesensamble desensamble = new frmDesensamble(formOpen);
                            break;
                        default:
                            break;
                    }
                }
                catch (Exception e)
                {
                    application.MessageBox(e.Message);
                }
            }
            else
            {
                if (pVal.MenuUID == Constants.Menu_Crear || pVal.MenuUID == Constants.Menu_Buscar)
                {
                    var mForm = application.Forms.ActiveForm;
                    Controles(mForm);
                }
            }

            try
            {
                // Control "Crear" de la barra de herramientas || Control "Buscar" de la barra de herramientas
                if (pVal.MenuUID == Constants.Menu_Crear || pVal.MenuUID == Constants.Menu_Buscar
                    || pVal.MenuUID == Constants.Registro_Datos_Anterior || pVal.MenuUID == Constants.Registro_Datos_Siguiente
                    || pVal.MenuUID == Constants.Primer_Registro_Datos || pVal.MenuUID == Constants.Ultimo_Registro_Datos)
                {
                    var mForm = application.Forms.ActiveForm;
                    if (formOpen.ContainsKey(mForm.UniqueID))
                        result = formOpen[mForm.UniqueID].HandleMenuDataEvents(pVal);
                }

                //Controles basados en el menu "Click derecho"
                if (pVal.MenuUID == Constants.Menu_AgregarLinea || pVal.MenuUID == Constants.Menu_EliminarLinea || pVal.MenuUID == Constants.Menu_Cancelar)
                {
                    if (pVal.BeforeAction)
                    {
                        var mForm = application.Forms.ActiveForm;
                        if (formOpen.ContainsKey(mForm.UniqueID))
                            result = formOpen[mForm.UniqueID].HandleMenuDataEvents(pVal);
                    }
                }
            }
            catch (Exception e)
            {
                application.MessageBox(e.Message);
            }
            BubbleEvent = result;
        }

        void Controles(SAPbouiCOM.Form mForm)
        {
            SAPbouiCOM.Matrix oMatrix;
            SAPbouiCOM.Item oEdtDate, oEdtWhs;
            SAPbouiCOM.Item btnOW, btnNS, btnAdd, btnSE, btnCE;
            string EDT_DATE = "edtDate", EDT_WHS = "edtWHS"; //EditTexts Porc
            string BTN_OK = "1", BTN_ORDFAB = "btnOrd", BTN_ENSAMBLAR = "btnExe", BTN_ADD = "btnAdd", BTN_SERIE = "btnSerie", BTN_FILE = "btnFile";//Buttons
            string MTX_MAIN = "mtxMain"; //Matrix

            switch (mForm.TypeEx)
            {
                case "CPRENS":
                    btnOW = (SAPbouiCOM.Item)mForm.Items.Item(BTN_ORDFAB);
                    btnNS = (SAPbouiCOM.Item)mForm.Items.Item(BTN_ENSAMBLAR);
                    btnAdd = (SAPbouiCOM.Item)mForm.Items.Item(BTN_ADD);
                    btnCE = (SAPbouiCOM.Item)mForm.Items.Item(BTN_FILE);
                    btnSE = (SAPbouiCOM.Item)mForm.Items.Item(BTN_SERIE);

                    switch (mForm.Mode)
                    {
                        case SAPbouiCOM.BoFormMode.fm_FIND_MODE:
                            btnAdd.Enabled = false;
                            btnCE.Enabled = false;
                            btnOW.Enabled = false;
                            btnNS.Enabled = false;
                            btnAdd.Enabled = false;
                            break;
                        case SAPbouiCOM.BoFormMode.fm_ADD_MODE:
                            btnAdd.Enabled = false;
                            btnCE.Enabled = true;
                            btnOW.Enabled = false;
                            btnNS.Enabled = false;
                            btnAdd.Enabled = false;
                            break;
                    }

                    break;
                case "CPRDES":
                    oMatrix = (SAPbouiCOM.Matrix)mForm.Items.Item(MTX_MAIN).Specific;
                    oEdtDate = (SAPbouiCOM.Item)mForm.Items.Item(EDT_DATE);
                    //oEdtWhs = (SAPbouiCOM.Item)mForm.Items.Item(EDT_WHS);
                    btnOW = (SAPbouiCOM.Item)mForm.Items.Item(BTN_ORDFAB);
                    btnNS = (SAPbouiCOM.Item)mForm.Items.Item(BTN_ENSAMBLAR);
                    btnAdd = (SAPbouiCOM.Item)mForm.Items.Item(BTN_ADD);
                    btnCE = (SAPbouiCOM.Item)mForm.Items.Item(BTN_FILE);
                    btnSE = (SAPbouiCOM.Item)mForm.Items.Item(BTN_SERIE);

                    switch (mForm.Mode)
                    {
                        case SAPbouiCOM.BoFormMode.fm_FIND_MODE:
                            oMatrix.Item.Enabled = true;
                            oEdtDate.Enabled = true;
                            btnAdd.Enabled = false;
                            btnCE.Enabled = true;
                            btnOW.Enabled = false;
                            btnNS.Enabled = false;
                            btnAdd.Enabled = false;
                            break;
                        case SAPbouiCOM.BoFormMode.fm_ADD_MODE:
                            oMatrix.Item.Enabled = true;
                            oEdtDate.Enabled = true;
                            btnAdd.Enabled = false;
                            btnCE.Enabled = true;
                            btnOW.Enabled = false;
                            btnNS.Enabled = false;
                            btnAdd.Enabled = false;
                            break;

                    }
                    break;
            }
        }

        void Application_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            try
            {
                BubbleEvent = true;

                if (formOpen.ContainsKey(BusinessObjectInfo.FormUID))
                {
                    BubbleEvent = formOpen[BusinessObjectInfo.FormUID].HandleFormDataEvents(BusinessObjectInfo);
                }
                else
                {
                    switch (BusinessObjectInfo.FormTypeEx)
                    {
                        //case FormName.CONTRATO:
                        //    BubbleEvent = new frmOCTR().HandleFormDataEvents(BusinessObjectInfo);
                        //    break;
                        //case FormName.LLAMADA_SERVICIO:
                        //    BubbleEvent = new frmOCSL().HandleFormDataEvents(BusinessObjectInfo);
                        //    break;
                        //default:
                        //    break;
                    }
                }
            }
            catch (Exception)
            {
                BubbleEvent = true;
            }
        }

        void Application_RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = formOpen[eventInfo.FormUID].HandleRightClickEvent(eventInfo);
        }

        //Creación de menú
        private void CrearMenu(System.Drawing.Bitmap imageBMP = null)
        {
            SAPbouiCOM.Form frmEps = application.Forms.GetFormByTypeAndCount(169, 1);
            frmEps.Freeze(true);
            try
            {
                application.StatusBar.SetText(Constants.PREFIX_MSG_ADDON + "Cargando opciones de menú", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_None);

                XmlDocument xmlMenu = new XmlDocument();
                xmlMenu.LoadXml(AddonProduccionEnsDes.Properties.Resources.Menu);
                application.LoadBatchActions(xmlMenu.InnerXml);
            }
            catch (Exception e)
            {
                application.StatusBar.SetText(Constants.PREFIX_MSG_ADDON + e.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                frmEps.Freeze(false);
                frmEps.Update();
            }
        }

        public static void AddForm(string UID, IForm newForm)
        {
            formOpen.Add(UID, newForm);
        }
    }
}