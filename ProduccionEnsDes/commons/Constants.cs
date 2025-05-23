using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AddonProduccionEnsDes.commons
{
    public class Constants
    {
        public const string PREFIX_MSG_ADDON = "[EXXIS - Prod. Ens/Des] ";
        public const string PREFIX_SP_HANA = "CALL ";
        public const string PREFIX_SP_SQL = "EXEC ";
        public const string SAPDATE_FORMAT = "yyyyMMdd";
        public const string HANADATE_FORMAT = "yyyy-MM-dd";
        public const string PERDATE_FORMAT = "MM-yyyy";
        public const string MAIN_CURR = "SOL";
        public const string ITEMS_CODE = "I";
        public const string SERV_CODE = "S";
        public const string ACTF_CODE = "A";


        //Botones de la barra de herramientas
        public const string Menu_Buscar = "1281";
        public const string Menu_Crear = "1282";
        public const string Menu_Context = "1280";
        public const string Menu_ViewLayout= "519";
        public const string Registro_Datos_Siguiente = "1288";
        public const string Registro_Datos_Anterior = "1289";
        public const string Primer_Registro_Datos = "1290";
        public const string Ultimo_Registro_Datos = "1291";
        public const string Actualizar_Registro = "1304";

        //Menú click derecho
        public const string Menu_AgregarLinea = "mnu_AgregarLinea";
        public const string Menu_EliminarLinea = "mnu_EliminarLinea";
        public const string Menu_AgregarLineaDescripcion = "Agregar fila";
        public const string Menu_EliminarLineaDescripcion = "Eliminar fila";
        public const string Menu_Cancelar = "1284";

        //Tipo Referencia Ruta
        public const int RutaBase = 0;
        public const int RutaOrigen = 1;
        public const int RutaDestino = 2;

        //Tipo Ruta
        public const string RT_EXTRUS = "EXT";
        public const string RT_IMPRES = "IMP";
        public const string RT_LAMINA = "LAM";
        public const string RT_SELLAD = "SEL";
        public const string RT_CORTE = "COR";
        public const string RT_HABILI = "HAB";
        public const string RT_REBOBI = "REB";
        public const string RT_SERV = "SER";
    }
}
