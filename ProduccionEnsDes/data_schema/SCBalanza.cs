using System.Collections.Generic;

namespace AddonProduccionEnsDes.data_schema
{
    class SCBalanza
    {
        #region _CABECERA_TABLA
        public const string TABLE_BALANZA = "EXP_BALANZA";
        public const string TABLE_BALANZA_DES = "EXP - Peso de Balanza";
        public const string TABLE_UNITALT = "EXP_UNIALT";
        public const string TABLE_UNITALT_DES = "EXP - Unidades Alternativas";
        public const string TABLE_USRRUT = "EXP_USRRUT";
        public const string TABLE_USRRUT_DES = "EXP - Usuario x Ruta";
        public const string TABLE_PARTENT= "EXP_PARTENT";
        public const string TABLE_PARTENT_DES = "EXP - Parte Entrega";
        public const string TABLE_ROLBAL = "EXP_ROLBAL";
        public const string TABLE_ROLBAL_DES = "EXP - Roles balanza";
        #endregion

        #region _CAMPOS
        public static List<CampoModel> getCamposTablas()
        {
            var myList = new List<CampoModel>();

            #region BALANZA
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_BALANZA,
                nombre_campo = "EXP_OrdFab",
                descrp_campo = "Orden de Fabricacion",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Numeric,
                tamano = 11
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_BALANZA,
                nombre_campo = "EXP_Etapa",
                descrp_campo = "Etapa",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_BALANZA,
                nombre_campo = "EXP_PesoNet",
                descrp_campo = "Peso Neto",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo=SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_BALANZA,
                nombre_campo = "EXP_Nro",
                descrp_campo = "Nro",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Numeric,
                tamano = 11
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_BALANZA,
                nombre_campo = "EXP_Fecha",
                descrp_campo = "Fecha",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Date
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_BALANZA,
                nombre_campo = "EXP_Hora",
                descrp_campo = "Hora",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Date,
                subtipo_campo=SAPbobsCOM.BoFldSubTypes.st_Time
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_BALANZA,
                nombre_campo = "EXP_Recibo",
                descrp_campo = "Nro de Recibo de Producción",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Numeric,
                tamano = 11
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_BALANZA,
                nombre_campo = "EXP_Recurso",
                descrp_campo = "Máquina",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_BALANZA,
                nombre_campo = "EXP_Art",
                descrp_campo = "Articulo",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_BALANZA,
                nombre_campo = "EXP_DscArt",
                descrp_campo = "Descripción",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_BALANZA,
                nombre_campo = "EXP_Und",
                descrp_campo = "Unidad",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_BALANZA,
                nombre_campo = "EXP_UndAlt",
                descrp_campo = "Unidad Alternativa",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_BALANZA,
                nombre_campo = "EXP_Ctd",
                descrp_campo = "Cantidad",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo=SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_BALANZA,
                nombre_campo = "EXP_PesoBrt",
                descrp_campo = "Peso Bruto",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_BALANZA,
                nombre_campo = "EXP_Tara",
                descrp_campo = "Tara",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_BALANZA,
                nombre_campo = "EXP_Tipo",
                descrp_campo = "Tipo",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 1,
                validValues = new string[] { "F", "B", "S","R", "D" },
                validDescription = new string[] { "Fardo", "Bobina", "Merma","Refile", "Devolucion" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_BALANZA,
                nombre_campo = "EXP_Medida",
                descrp_campo = "Medida Ruta",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano=100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_USRRUT,
                nombre_campo = "EXP_USER",
                descrp_campo = "Usuario",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_USRRUT,
                nombre_campo = "EXP_RUTA",
                descrp_campo = "Ruta",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 10
            });
            //myList.Add(new CampoBean()
            //{
            //    nombre_tabla = TABLE_BALANZA,
            //    nombre_campo = "EXP_LOTE",
            //    descrp_campo = "Lote",
            //    tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
            //    tamano = 36
            //});
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_BALANZA,
                nombre_campo = "EXP_Material",
                descrp_campo = "Tipo Material",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_BALANZA,
                nombre_campo = "EXP_ParEntrega",
                descrp_campo = "Parte Entrega",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_BALANZA,
                nombre_campo = "EXP_Obs",
                descrp_campo = "Con Observaciones",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 1,
                validValues = new string[] { "Y", "N" },
                validDescription = new string[] { "Si", "No" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_BALANZA,
                nombre_campo = "EXP_StaEnt",
                descrp_campo = "Estado entrega",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 1,
                validValues = new string[] { "0", "1", "2", "3", "X" },
                validDescription = new string[] { "Pendiente", "Valida Parte", "Valida Calidad", "Entregado", "Anulado" }
            });
             myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_BALANZA,
                nombre_campo = "EXP_AprCal",
                descrp_campo = "Aprobador Calidad",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 20
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_BALANZA,
                nombre_campo = "EXP_MedidaMan",
                descrp_campo = "Medida Manual",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_BALANZA,
                nombre_campo = "EXP_EspesorMan",
                descrp_campo = "Espesor Manual",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            }); myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_BALANZA,
                nombre_campo = "EXP_DestinoMan",
                descrp_campo = "Destino Manual",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            #endregion
            #region ParteEntrega
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_PARTENT,
                nombre_campo = "EXP_BAL",
                descrp_campo = "Codigo Balanza",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_PARTENT,
                nombre_campo = "EXP_Fecha",
                descrp_campo = "Fecha",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Date
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_PARTENT,
                nombre_campo = "EXP_Hora",
                descrp_campo = "Hora",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Date,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Time
            });
        
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_PARTENT,
                nombre_campo = "EXP_Tipo",
                descrp_campo = "Tipo",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 1,
                validValues = new string[] { "F", "B", "S", "R", "D" },
                validDescription = new string[] { "Fardo", "Bobina", "Merma", "Refile", "Devolucion" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_PARTENT,
                nombre_campo = "EXP_USER",
                descrp_campo = "Usuario",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            #endregion
            #region Rol x Balanza
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_ROLBAL,
                nombre_campo = "EXP_Pesar",
                descrp_campo = "Pesar Balanza",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 1,
                validValues = new string[] { "Y", "N" },
                validDescription = new string[] { "Si", "No" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_ROLBAL,
                nombre_campo = "EXP_Eliminar",
                descrp_campo = "Eliminar Balanza",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 1,
                validValues = new string[] { "Y", "N" },
                validDescription = new string[] { "Si", "No" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_ROLBAL,
                nombre_campo = "EXP_ElimPE",
                descrp_campo = "Elim. Balanza con PE",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 1,
                validValues = new string[] { "Y", "N" },
                validDescription = new string[] { "Si", "No" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_ROLBAL,
                nombre_campo = "EXP_Parte",
                descrp_campo = "Parte Entrega",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 1,
                validValues = new string[] { "Y", "N" },
                validDescription = new string[] { "Si", "No" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_ROLBAL,
                nombre_campo = "EXP_EnvSAP",
                descrp_campo = "Entregar SAP",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 1,
                validValues = new string[] { "Y", "N" },
                validDescription = new string[] { "Si", "No" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_ROLBAL,
                nombre_campo = "EXP_Calidad",
                descrp_campo = "Aprobar Calidad",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 1,
                validValues = new string[] { "Y", "N" },
                validDescription = new string[] { "Si", "No" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_ROLBAL,
                nombre_campo = "EXP_EnvCtrH",
                descrp_campo = "Ctrll Hora SaP ",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 1,
                validValues = new string[] { "Y", "N" },
                validDescription = new string[] { "Si", "No" }
            });
            #endregion
            return myList;
        }
        #endregion
    }
}
