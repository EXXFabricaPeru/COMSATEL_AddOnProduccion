using System.Collections.Generic;

namespace AddonProduccionEnsDes.data_schema
{
    class SCEnsamble
    {
        #region _CABECERA_TABLA
        public const string TABLE_CABE = "EXC_ENSA", TABLE_CABE_DES = "EXC - Ensamble";
        public const string TABLE_DET1 = "EXC_ENS1", TABLE_DET1_DES = "EXC - Det.Ensamble";
        #endregion

        #region _CAMPOS
        public static List<CampoModel> getCamposTablas()
        {
            var myList = new List<CampoModel>();

            #region CABECERA
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_CABE,
                nombre_campo = "EXC_FEPR",
                descrp_campo = "Fecha Produccion",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Date
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_CABE,
                nombre_campo = "EXC_ESTA",
                descrp_campo = "Estado",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 1,
                validValues = new string[] { "O", "P", "F" },
                validDescription = new string[] { "Abierto", "En Producción", "Finalizado" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_CABE,
                nombre_campo = "EXC_ALMA",
                descrp_campo = "Almacen Ingreso",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 8
            });


            #endregion
            #region DETALLE
            myList.Add(new CampoModel() { nombre_tabla = TABLE_DET1, nombre_campo = "EXC_CPRO", descrp_campo = "Cod Producir", tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha, tamano = 50 });
            myList.Add(new CampoModel() { nombre_tabla = TABLE_DET1, nombre_campo = "EXC_DPRO", descrp_campo = "Desc Producir", tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha, tamano = 200 });
            myList.Add(new CampoModel() { nombre_tabla = TABLE_DET1, nombre_campo = "EXC_CEQP", descrp_campo = "Cod Equipo", tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha, tamano = 50 });
            myList.Add(new CampoModel() { nombre_tabla = TABLE_DET1, nombre_campo = "EXC_DEQP", descrp_campo = "Desc Equipo", tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha, tamano = 200 });
            myList.Add(new CampoModel() { nombre_tabla = TABLE_DET1, nombre_campo = "EXC_SEQP", descrp_campo = "Ser Equipo", tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha, tamano = 36 });
            myList.Add(new CampoModel() { nombre_tabla = TABLE_DET1, nombre_campo = "EXC_IMEIEQP", descrp_campo = "IMEI Equipo", tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha, tamano = 20 });
            myList.Add(new CampoModel() { nombre_tabla = TABLE_DET1, nombre_campo = "EXC_MARC", descrp_campo = "Marca", tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha, tamano = 30 });
            myList.Add(new CampoModel() { nombre_tabla = TABLE_DET1, nombre_campo = "EXC_MODE", descrp_campo = "Modelo", tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha, tamano = 30 });
            myList.Add(new CampoModel() { nombre_tabla = TABLE_DET1, nombre_campo = "EXC_CCHI", descrp_campo = "Cod Chip", tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha, tamano = 50 });
            myList.Add(new CampoModel() { nombre_tabla = TABLE_DET1, nombre_campo = "EXC_DCHI", descrp_campo = "Desc Chip", tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha, tamano = 200 });
            myList.Add(new CampoModel() { nombre_tabla = TABLE_DET1, nombre_campo = "EXC_SCHI", descrp_campo = "Serie Chip", tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha, tamano = 36 });
            myList.Add(new CampoModel() { nombre_tabla = TABLE_DET1, nombre_campo = "EXC_IMEI", descrp_campo = "IMEI", tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha, tamano = 20 });
            myList.Add(new CampoModel() { nombre_tabla = TABLE_DET1, nombre_campo = "EXC_CACC1", descrp_campo = "Cod Accesorio 1", tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha, tamano = 50 });
            myList.Add(new CampoModel() { nombre_tabla = TABLE_DET1, nombre_campo = "EXC_DACC1", descrp_campo = "Desc Accesorio", tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha, tamano = 200 });
            myList.Add(new CampoModel() { nombre_tabla = TABLE_DET1, nombre_campo = "EXC_CACC2", descrp_campo = "Cod Accesorio 2", tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha, tamano = 50 });
            myList.Add(new CampoModel() { nombre_tabla = TABLE_DET1, nombre_campo = "EXC_DACC2", descrp_campo = "Desc Accesorio", tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha, tamano = 200 }); 
            myList.Add(new CampoModel() { nombre_tabla = TABLE_DET1, nombre_campo = "EXC_OPER", descrp_campo = "Operador", tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha, tamano = 10 });
            myList.Add(new CampoModel() { nombre_tabla = TABLE_DET1, nombre_campo = "EXC_VERS", descrp_campo = "Version", tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha, tamano = 12 });
            myList.Add(new CampoModel() { nombre_tabla = TABLE_DET1, nombre_campo = "EXC_ARCH", descrp_campo = "Archivo Config", tipo_campo = SAPbobsCOM.BoFieldTypes.db_Memo });
            myList.Add(new CampoModel() { nombre_tabla = TABLE_DET1, nombre_campo = "EXC_PROP", descrp_campo = "Producido por", tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha, tamano = 50 });
            myList.Add(new CampoModel() { nombre_tabla = TABLE_DET1, nombre_campo = "EXC_NSER", descrp_campo = "Nro Serie", tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha, tamano = 36 });
            myList.Add(new CampoModel() { nombre_tabla = TABLE_DET1, nombre_campo = "EXC_ORDT", descrp_campo = "Orden Trabajo", tipo_campo = SAPbobsCOM.BoFieldTypes.db_Numeric, tamano = 11 });
            //nuevos campos
            myList.Add(new CampoModel() { nombre_tabla = TABLE_DET1, nombre_campo = "EXC_SERV", descrp_campo = "Servicio", tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha, tamano = 30 });
            myList.Add(new CampoModel() { nombre_tabla = TABLE_DET1, nombre_campo = "EXC_FWAR", descrp_campo = "FirmWare", tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha, tamano = 50 });
            myList.Add(new CampoModel() { nombre_tabla = TABLE_DET1, nombre_campo = "EXC_FOTA", descrp_campo = "FOTA", tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha, tamano = 1, validValues =  new string[] { "Y", "N"}, validDescription = new string[] { "Si", "No"} });
            myList.Add(new CampoModel() { nombre_tabla = TABLE_DET1, nombre_campo = "EXC_NSLO", descrp_campo = "Nro Serie Lote", tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha, tamano = 50 });
            

            myList.Add(new CampoModel() { nombre_tabla = TABLE_DET1, nombre_campo = "EXC_TIIP", descrp_campo = "Tipo IP", tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha, tamano = 20 });
            myList.Add(new CampoModel() { nombre_tabla = TABLE_DET1, nombre_campo = "EXC_NRIP", descrp_campo = "Nro IP", tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha, tamano = 20 });
            myList.Add(new CampoModel() { nombre_tabla = TABLE_DET1, nombre_campo = "EXC_PQDA", descrp_campo = "Paquete de datos", tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha, tamano = 30 });
            myList.Add(new CampoModel() { nombre_tabla = TABLE_DET1, nombre_campo = "EXC_LINE", descrp_campo = "Linea", tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha, tamano = 30 });
            myList.Add(new CampoModel() { nombre_tabla = TABLE_DET1, nombre_campo = "EXC_LTEL", descrp_campo = "Linea telefonica", tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha, tamano = 20 });
            myList.Add(new CampoModel() { nombre_tabla = TABLE_DET1, nombre_campo = "EXC_DAPN", descrp_campo = "APN", tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha, tamano = 30 });
            myList.Add(new CampoModel() { nombre_tabla = TABLE_DET1, nombre_campo = "EXC_SIMC", descrp_campo = "SIMCARD", tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha, tamano = 30 });
            myList.Add(new CampoModel() { nombre_tabla = TABLE_DET1, nombre_campo = "EXC_PROT", descrp_campo = "Protocolo", tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha, tamano = 30 });

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
                name = "EXC_ENSAMBLE",
                tableName = TABLE_CABE,
                canCancel = SAPbobsCOM.BoYesNoEnum.tNO,
                canClose = SAPbobsCOM.BoYesNoEnum.tYES,
                canDelete = SAPbobsCOM.BoYesNoEnum.tNO,
                childTables = new string[] { TABLE_DET1 },
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
