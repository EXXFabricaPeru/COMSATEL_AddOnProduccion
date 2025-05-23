using System.Collections.Generic;

namespace AddonProduccionEnsDes.data_schema
{
    class SCPeriodica
    {
        #region _CABECERA_TABLA
        public const string TABLE_CABE = "EXC_RECUR", TABLE_CABE_DES = "EXC - Recurrente";
        public const string TABLE_DET1 = "EXC_RECUR1", TABLE_DET1_DES = "EXC - Det.Recurrente";
        #endregion

        #region _CAMPOS
        public static List<CampoModel> getCamposTablas()
        {
            var myList = new List<CampoModel>();

            #region CABECERA
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_CABE,
                nombre_campo = "EXC_IDENT",
                descrp_campo = "Identificador",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 20
            });
            
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_CABE,
                nombre_campo = "EXC_CONTID",
                descrp_campo = "ID Contrato",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Numeric,
                tamano = 11
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_CABE,
                nombre_campo = "EXC_ORDREF",
                descrp_campo = "ID Orden Referencia",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Numeric,
                
                tamano = 11
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_CABE,
                nombre_campo = "EXC_ORDBAS",
                descrp_campo = "ID Orden Base",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Numeric,
                tamano = 11
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_CABE,
                nombre_campo = "EXC_PERIFAC",
                descrp_campo = "Periodo",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50,
                tablaVinculada="EXC_PERIFAC"
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_CABE,
                nombre_campo = "EXC_DIAREP",
                descrp_campo = "Dia Repeticion",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Numeric,
                tamano = 2
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_CABE,
                nombre_campo = "EXC_FECINI",
                descrp_campo = "Fecha Inicio",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Date
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_CABE,
                nombre_campo = "EXC_FECFIN",
                descrp_campo = "Fecha Fin",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Date
            });

     

            #endregion
            #region DETALLE
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_DET1,
                nombre_campo = "EXC_ORDREC",
                descrp_campo = "ID Orden Rec.",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Numeric,
                tamano = 11
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_DET1,
                nombre_campo = "EXC_ESTADO",
                descrp_campo = "Estado Ejecucion",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 1,
                validValues = new string[] { "P", "E" },
                validDescription = new string[] { "Pendiente", "Ejecutado" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_DET1,
                nombre_campo = "EXC_FECINI",
                descrp_campo = "Fecha Ejecucion",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Date
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_DET1,
                nombre_campo = "EXC_TOTDOC",
                descrp_campo = "Total documento",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo=SAPbobsCOM.BoFldSubTypes.st_Price
            });
            #endregion
            return myList;
        }

        #endregion
        #region _OBJETO
        public static ObjetoModel getObjeto()
        {
            var myObj = new ObjetoModel
            {
                code = TABLE_CABE,
                name = "EXC_RECURRENTES",
                tableName = TABLE_CABE,
                canCancel = SAPbobsCOM.BoYesNoEnum.tNO,
                canClose = SAPbobsCOM.BoYesNoEnum.tNO,
                canDelete = SAPbobsCOM.BoYesNoEnum.tNO,
                childTables = new string[] { TABLE_DET1},
                canCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO,
                canFind = SAPbobsCOM.BoYesNoEnum.tYES,
                canLog = SAPbobsCOM.BoYesNoEnum.tNO,
                objectType = SAPbobsCOM.BoUDOObjType.boud_Document,
                enableEnhancedForm = SAPbobsCOM.BoYesNoEnum.tNO,
                rebuildEnhancedForm = SAPbobsCOM.BoYesNoEnum.tNO
            };
            return myObj;
        }
        #endregion


    }
}
