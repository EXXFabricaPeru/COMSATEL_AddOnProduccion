using System.Collections.Generic;

namespace AddonProduccionEnsDes.data_schema
{
    class SCCtrHora
    {
        #region _CABECERA_TABLA
        public const string TABLE_CTRLHORA = "EXP_CTRHRA";
        public const string TABLE_CTRLHORA_DES = "EXP - Control Tiempos";
        public const string TABLE_PARADA = "EXP_PARADA";
        public const string TABLE_PARADA_DES = "EXP - Parada Control Tiempos";
        public const string TABLE_TIPPAR = "EXP_TIPPAR";
        public const string TABLE_TIPPAR_DES = "EXP - Tipo Parada";
        #endregion

        #region _CAMPOS
        public static List<CampoModel> getCamposTablas()
        {
            var myList = new List<CampoModel>();

            #region CTRLHORA
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_CTRLHORA,
                nombre_campo = "EXP_OrdFab",
                descrp_campo = "Orden de Fabricacion",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Numeric,
                tamano = 11
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_CTRLHORA,
                nombre_campo = "EXP_Etapa",
                descrp_campo = "Etapa",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_CTRLHORA,
                nombre_campo = "EXP_Recurso",
                descrp_campo = "Máquina",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });

            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_CTRLHORA,
                nombre_campo = "EXP_Tipo",
                descrp_campo = "Tipo",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 2,
                validValues = new string[] { "01", "02" },
                validDescription = new string[] { "Inicio", "Fin" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_CTRLHORA,
                nombre_campo = "EXP_Fecha",
                descrp_campo = "Fecha",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Date
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_CTRLHORA,
                nombre_campo = "EXP_Hora",
                descrp_campo = "Hora",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Date,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Time
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_CTRLHORA,
                nombre_campo = "EXP_CHOri",
                descrp_campo = "Origen",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Numeric,
                tamano = 11
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_CTRLHORA,
                nombre_campo = "EXP_Comp",
                descrp_campo = "Nro Comprobante",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Numeric,
                tamano = 11
            });
            #endregion
            #region PARADA
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_PARADA,
                nombre_campo = "EXP_OrdFab",
                descrp_campo = "Orden de Fabricacion",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Numeric,
                tamano = 11
            });

            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_PARADA,
                nombre_campo = "EXP_Etapa",
                descrp_campo = "Etapa",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_PARADA,
                nombre_campo = "EXP_Recurso",
                descrp_campo = "Maquina",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_PARADA,
                nombre_campo = "EXP_TipPara",
                descrp_campo = "Tipo de Para",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_PARADA,
                nombre_campo = "EXP_TipParaDesc",
                descrp_campo = "Tipo Para Desc",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_PARADA,
                nombre_campo = "EXP_MotPara",
                descrp_campo = "Motivo de Para",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 70
            });

            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_PARADA,
                nombre_campo = "EXP_Tipo",
                descrp_campo = "Tipo",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 2,
                validValues = new string[] { "01", "02" },
                validDescription = new string[] { "Inicio de Para", "Fin de Para" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_PARADA,
                nombre_campo = "EXP_Fecha",
                descrp_campo = "Fecha",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Date
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_PARADA,
                nombre_campo = "EXP_Hora",
                descrp_campo = "Hora",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Date,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Time
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_PARADA,
                nombre_campo = "EXP_CHOri",
                descrp_campo = "Origen",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Numeric,
                tamano = 11
            });
            #endregion
            #region TIPOPARADA
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_TIPPAR,
                nombre_campo = "EXP_Suma",
                descrp_campo = "Parada considerada",
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
