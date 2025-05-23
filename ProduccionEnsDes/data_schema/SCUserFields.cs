using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AddonProduccionEnsDes.data_schema
{
    public class SCUserFields
    {
        #region _CABECERA_TABLA
        public const string TABLE_SERIE = "OSRN";

        #endregion

        #region _COLUMNAS
        public static List<CampoModel> getCamposUsuario()
        {
            var myList = new List<CampoModel>();
            #region Recursos
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_SERIE,
                nombre_campo = "EXC_PRODPOR",
                descrp_campo = "Producido Por",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_SERIE,
                nombre_campo = "EXC_VERSION",
                descrp_campo = "Version",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 12
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_SERIE,
                nombre_campo = "EXC_CACC1",
                descrp_campo = "Accesorio",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 12
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_SERIE,
                nombre_campo = "EXC_CACC2",
                descrp_campo = "Accesorio 2",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 12
            });
            #endregion
            return myList;
        }
        #endregion
    }
}
