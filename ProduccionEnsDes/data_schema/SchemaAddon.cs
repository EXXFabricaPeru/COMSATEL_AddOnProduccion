using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AddonProduccionEnsDes.data_schema
{
    public class SchemaAddon
    {
        #region TABLAS_GENERICAS
        public static Dictionary<string, string> tablesGeneric()
        {
            var tables = new Dictionary<string, string>();
            //tables.Add(SCLinProd.TABLE_CABE, SCLinProd.TABLE_CABE_DES);
            //tables.Add(SCCtrHora.TABLE_CTRLHORA, SCCtrHora.TABLE_CTRLHORA_DES);
            //tables.Add(SCCtrHora.TABLE_PARADA, SCCtrHora.TABLE_PARADA_DES);
            //tables.Add(SCCtrHora.TABLE_TIPPAR, SCCtrHora.TABLE_TIPPAR_DES);
            //tables.Add(SCBalanza.TABLE_BALANZA, SCBalanza.TABLE_BALANZA_DES);
            //tables.Add(SCBalanza.TABLE_USRRUT, SCBalanza.TABLE_USRRUT_DES);
            //tables.Add(SCBalanza.TABLE_PARTENT, SCBalanza.TABLE_PARTENT_DES);
            //tables.Add(SCBalanza.TABLE_ROLBAL, SCBalanza.TABLE_ROLBAL_DES);
            tables.Add(SCConfig.TABLE_SERIEPROD, SCConfig.TABLE_SERIEPROD_DES);
            //tables.Add(SCConfig.TABLE_ALMACENFORM, SCConfig.TABLE_ALMACENFORM_DES);

            return tables;
        }
        public static Dictionary<string, string> tablesGenericNoAuto()
        {
            var tables = new Dictionary<string, string>();
            //tables.Add(SCBalanza.TABLE_UNITALT, SCBalanza.TABLE_UNITALT_DES);
            //tables.Add(SCConfig.TABLE_CONVERSION, SCConfig.TABLE_CONVERSION_DES);
            //tables.Add(SCConfig.TABLE_TINTA, SCConfig.TABLE_TINTA_DES);

            return tables;
        }
        #endregion
        #region TABLAS_DATOS_MAESTROS
        //Cabeceras
        public static Dictionary<string, string> tablesMasterH()
        {
            var tables = new Dictionary<string, string>();
            //tables.Add(SCConfig.TABLE_CABE, SCConfig.TABLE_CABE_DES);
            //tables.Add(SCFormulacion.TABLE_CABE, SCFormulacion.TABLE_CABE_DES);
            return tables;
        }

        //Detalles
        public static Dictionary<string, string> tablesMasterD()
        {
            var tables = new Dictionary<string, string>();
            //tables.Add(SCFormulacion.TABLE_FRRUTA, SCFormulacion.TABLE_FRRUTA_DES);
            //tables.Add(SCFormulacion.TABLE_FORMUL, SCFormulacion.TABLE_FORMUL_DES);
            //tables.Add(SCFormulacion.TABLE_SUBPRD, SCFormulacion.TABLE_SUBPRD_DES);
            //tables.Add(SCFormulacion.TABLE_INDUCT, SCFormulacion.TABLE_INDUCT_DES);
            //tables.Add(SCFormulacion.TABLE_FREXTR, SCFormulacion.TABLE_FREXTR_DES);
            //tables.Add(SCFormulacion.TABLE_FRIMPR, SCFormulacion.TABLE_FRIMPR_DES);
            //tables.Add(SCFormulacion.TABLE_LAMINA, SCFormulacion.TABLE_LAMINA_DES);
            //tables.Add(SCFormulacion.TABLE_FRSELA, SCFormulacion.TABLE_FRSELA_DES);
            //tables.Add(SCFormulacion.TABLE_FRCORT, SCFormulacion.TABLE_FRCORT_DES);
            //tables.Add(SCFormulacion.TABLE_FRHABI, SCFormulacion.TABLE_FRHABI_DES);
            //tables.Add(SCFormulacion.TABLE_FRREBO, SCFormulacion.TABLE_FRREBO_DES);
            //tables.Add(SCFormulacion.TABLE_FRSERV, SCFormulacion.TABLE_FRSERV_DES);

            return tables;
        }
        #endregion
        #region TABLAS_DOCUMENTOS
        //Cabeceras
        public static Dictionary<string, string> tablesDocsH()
        {
            var tables = new Dictionary<string, string>();
            tables.Add(SCDesensamble.TABLE_CABE, SCDesensamble.TABLE_CABE_DES);
            tables.Add(SCEnsamble.TABLE_CABE, SCEnsamble.TABLE_CABE_DES);

            return tables;
        }

        //Detalles
        public static Dictionary<string, string> tablesDocsD()
        {
            var tables = new Dictionary<string, string>();
            tables.Add(SCDesensamble.TABLE_DET1, SCDesensamble.TABLE_DET1_DES);
            tables.Add(SCEnsamble.TABLE_DET1, SCEnsamble.TABLE_DET1_DES);

            return tables;
        }
        #endregion
        public static List<CampoModel> camposTB()
        {
            var campos = new List<CampoModel>();
            campos.AddRange(SCUserFields.getCamposUsuario());
            campos.AddRange(SCEnsamble.getCamposTablas());
            campos.AddRange(SCDesensamble.getCamposTablas());
            campos.AddRange(SCConfig.getCamposTablas());
            //campos.AddRange(SCPeriodica.getCamposTablas());
            //campos.AddRange(SCFormulacion.getCamposTablas());
            //campos.AddRange(SCBalanza.getCamposTablas());
            //campos.AddRange(SCCtrHora.getCamposTablas());
            //campos.AddRange(SCConfig.getCamposTablas());
            return campos;
        }

        public static List<ObjetoModel> objetosADDON()
        {
            var objects = new List<ObjetoModel>();
            objects.Add(SCEnsamble.getObjeto());
            objects.Add(SCDesensamble.getObjeto());
            return objects;
        }
    }
}
