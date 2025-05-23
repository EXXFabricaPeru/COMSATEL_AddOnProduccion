using System.Collections.Generic;

namespace AddonProduccionEnsDes.data_schema
{
    class SCConfig
    {
        #region _CABECERA_TABLA
        public const string TABLE_CABE = "EXC_CFG_PED";
        public const string TABLE_CABE_DES = "EXC - Cfg. Prod. EnsDes";
        public const string TABLE_SERIEPROD = "EXC_SERIEPROD";
        public const string TABLE_SERIEPROD_DES = "EXC - Cor. Serie Prod.";
        #endregion

        #region _CAMPOS
        public static List<CampoModel> getCamposTablas()
        {
            var myList = new List<CampoModel>();

            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_SERIEPROD,
                nombre_campo = "EXC_MARCA",
                descrp_campo = "Marca",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano=30
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_SERIEPROD,
                nombre_campo = "EXC_MODELO",
                descrp_campo = "Modelo",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 30
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_SERIEPROD,
                nombre_campo = "EXC_PREFIJO",
                descrp_campo = "Prefijo",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 10
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_SERIEPROD,
                nombre_campo = "EXC_CORR",
                descrp_campo = "Correlativo",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Numeric,
                tamano = 5
            });
            return myList;
        }

        #endregion
        #region _OBJETO

        public static ObjetoModel getObjeto()
        {
            var myObj = new ObjetoModel
            {
                code = TABLE_CABE,
                name = "CONFIG_RUTA_LISTMAT",
                tableName = TABLE_CABE,
                canCancel = SAPbobsCOM.BoYesNoEnum.tNO,
                canClose = SAPbobsCOM.BoYesNoEnum.tNO,
                canDelete = SAPbobsCOM.BoYesNoEnum.tNO,
                childTables = new string[] {  },
                canCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO,
                canFind = SAPbobsCOM.BoYesNoEnum.tNO,
                canLog = SAPbobsCOM.BoYesNoEnum.tNO,
                objectType = SAPbobsCOM.BoUDOObjType.boud_MasterData,
                manageSeries = SAPbobsCOM.BoYesNoEnum.tNO,
                enableEnhancedForm = SAPbobsCOM.BoYesNoEnum.tNO,
                rebuildEnhancedForm = SAPbobsCOM.BoYesNoEnum.tNO
            };
            return myObj;
        }

        #endregion

    }
}
