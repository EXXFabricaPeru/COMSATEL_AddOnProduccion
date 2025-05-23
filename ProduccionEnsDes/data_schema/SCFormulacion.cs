using System.Collections.Generic;

namespace AddonProduccionEnsDes.data_schema
{
    class SCFormulacion
    {
        #region _CABECERA_TABLA
        public const string TABLE_CABE = "EXP_OFRM", TABLE_CABE_DES = "EXP - Formulacion";
        public const string TABLE_FRRUTA = "EXP_FFRUTA", TABLE_FRRUTA_DES = "EXP - Form. Ruta";
        public const string TABLE_FORMUL = "EXP_FORMUL", TABLE_FORMUL_DES = "EXP - Formula";
        public const string TABLE_SUBPRD = "EXP_SUBPRD", TABLE_SUBPRD_DES = "EXP - Subproducto";
        public const string TABLE_INDUCT = "EXP_INDUCT", TABLE_INDUCT_DES = "EXP - Inductor";
        public const string TABLE_FREXTR = "EXP_FREXTR", TABLE_FREXTR_DES = "EXP - Form. Extrusión";
        public const string TABLE_FRIMPR = "EXP_FRIMPR", TABLE_FRIMPR_DES = "EXP - Form. Impresión";
        public const string TABLE_LAMINA = "EXP_LAMINA", TABLE_LAMINA_DES = "EXP - Form. Laminado";
        public const string TABLE_FRSELA = "EXP_FRSELA", TABLE_FRSELA_DES = "EXP - Form.  Sellado";
        public const string TABLE_FRCORT = "EXP_FRCORT", TABLE_FRCORT_DES = "EXP - Form. Corte";
        public const string TABLE_FRHABI = "EXP_FRHABI", TABLE_FRHABI_DES = "EXP - Form. Habilitado";
        public const string TABLE_FRREBO = "EXP_FRREBO", TABLE_FRREBO_DES = "EXP - Form. Rebobinado";
        public const string TABLE_FRSERV = "EXP_FRSERV", TABLE_FRSERV_DES = "EXP - Form. Servicio";
        #endregion

        #region _CAMPOS
        public static List<CampoModel> getCamposTablas()
        {
            var myList = new List<CampoModel>();

            #region Cabecera_OFRM
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_CABE,
                nombre_campo = "EXP_QTT",
                descrp_campo = "Cantidad",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });

            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_CABE,
                nombre_campo = "EXP_PRCAP1",
                descrp_campo = "Porcentaje A",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Percentage
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_CABE,
                nombre_campo = "EXP_PRCAP2",
                descrp_campo = "Porcentaje B",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Percentage
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_CABE,
                nombre_campo = "EXP_PRCAP3",
                descrp_campo = "Porcentaje C",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Percentage
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_CABE,
                nombre_campo = "EXP_PRCAP4",
                descrp_campo = "Porcentaje D",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Percentage
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_CABE,
                nombre_campo = "EXP_PRCAP5",
                descrp_campo = "Porcentaje E",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Percentage
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_CABE,
                nombre_campo = "EXP_TIPO",
                descrp_campo = "Tipo Formula",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 2,
                validValues = new string[] { "PK", "CM" },
                validDescription = new string[] { "Peso Kilo", "Cantidad MLL" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_CABE,
                nombre_campo = "EXP_ARTI",
                descrp_campo = "Articulo Referencia",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 20
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_CABE,
                nombre_campo = "EXP_ARTD",
                descrp_campo = "Descripcion Art Ref",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            #endregion
            #region FRRUTA 
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRRUTA,
                nombre_campo = "EXP_CODRUT",
                descrp_campo = "Codigo",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
                //validValues = new string[] { "EXT", "IMP", "LAM", "COR", "SEL", "ENT" },
                //validDescription = new string[] { "Extrusora", "Impresion", "Laminado", "Corte", "Sellado", "Entrega" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRRUTA,
                nombre_campo = "EXP_NOMRUT",
                descrp_campo = "Descripcion",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRRUTA,
                nombre_campo = "EXP_SUBPRD",
                descrp_campo = "Subproducto",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRRUTA,
                nombre_campo = "EXP_SCRAP",
                descrp_campo = "Merma",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRRUTA,
                nombre_campo = "EXP_REFILE",
                descrp_campo = "Refile",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            #endregion
            #region FORMUL
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FORMUL,
                nombre_campo = "EXP_CODRUT",
                descrp_campo = "Codigo",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FORMUL,
                nombre_campo = "EXP_CODART",
                descrp_campo = "Artículo",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
                //TODO OITM
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FORMUL,
                nombre_campo = "EXP_DESCRP",
                descrp_campo = "Descripción",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 200
            });

            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FORMUL,
                nombre_campo = "EXP_CAPA1",
                descrp_campo = "A - KG",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FORMUL,
                nombre_campo = "EXP_PORC1",
                descrp_campo = "A - %",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Percentage
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FORMUL,
                nombre_campo = "EXP_CAPA2",
                descrp_campo = "B - KG",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FORMUL,
                nombre_campo = "EXP_CAPA3",
                descrp_campo = "C - KG",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FORMUL,
                nombre_campo = "EXP_CAPA4",
                descrp_campo = "D - KG",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FORMUL,
                nombre_campo = "EXP_CAPA5",
                descrp_campo = "E - KG",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });

            #endregion
            #region INDUCT 
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_INDUCT,
                nombre_campo = "EXP_CODRUT",
                descrp_campo = "Codigo",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_INDUCT,
                nombre_campo = "EXP_INDCTR",
                descrp_campo = "Inductor",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_INDUCT,
                nombre_campo = "EXP_STATUS",
                descrp_campo = "Activo",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 1,
                validValues = new string[] { "Y", "N" },
                validDescription = new string[] { "Si", "No" }
            });
            #endregion
            #region FREXTR
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FREXTR,
                nombre_campo = "EXP_CODRUTDE",
                descrp_campo = "Cod. Ruta Dest",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FREXTR,
                nombre_campo = "EXP_NOMRUTDE",
                descrp_campo = "Ruta Destino",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FREXTR,
                nombre_campo = "EXP_RECMAQ",
                descrp_campo = "Maquina Impresión",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FREXTR,
                nombre_campo = "EXP_TPOPRE",
                descrp_campo = "Preparación ",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FREXTR,
                nombre_campo = "EXP_ANCHO",
                descrp_campo = "Ancho",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FREXTR,
                nombre_campo = "EXP_UNANC",
                descrp_campo = "Unidad Espesor",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FREXTR,
                nombre_campo = "EXP_PESBOB",
                descrp_campo = "Peso Bobina",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FREXTR,
                nombre_campo = "EXP_MANG",
                descrp_campo = "Manga",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 1,
                validValues = new string[] { "N", "C", "1", "2" },
                validDescription = new string[] { "Ninguno", "Cerrado", "Abierto 1 lado", "Abierto 2 lados" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FREXTR,
                nombre_campo = "EXP_LAMI",
                descrp_campo = "Lamina",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 1,
                validValues = new string[] { "Y", "N" },
                validDescription = new string[] { "Si", "No" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FREXTR,
                nombre_campo = "EXP_TRAT",
                descrp_campo = "Tratada",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 1,
                validValues = new string[] { "N", "T", "S" },
                validDescription = new string[] { "Ninguna", "Total", "Sectorizado" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FREXTR,
                nombre_campo = "EXP_MICR",
                descrp_campo = "Microperforado",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 1,
                validValues = new string[] { "Y", "N" },
                validDescription = new string[] { "Si", "No" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FREXTR,
                nombre_campo = "EXP_FUEL",
                descrp_campo = "Fuelle",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 1,
                validValues = new string[] { "Y", "N" },
                validDescription = new string[] { "Si", "No" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FREXTR,
                nombre_campo = "EXP_MATE",
                descrp_campo = "Material",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FREXTR,
                nombre_campo = "EXP_CRIS",
                descrp_campo = "Cristal",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 1,
                validValues = new string[] { "Y", "N" },
                validDescription = new string[] { "Si", "No" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FREXTR,
                nombre_campo = "EXP_RECP",
                descrp_campo = "Recuperado",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 1,
                validValues = new string[] { "Y", "N" },
                validDescription = new string[] { "Si", "No" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FREXTR,
                nombre_campo = "EXP_COLOR",
                descrp_campo = "Color",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FREXTR,
                nombre_campo = "EXP_CLOS",
                descrp_campo = "Cerrada",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 1,
                validValues = new string[] { "Y", "N" },
                validDescription = new string[] { "Si", "No" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FREXTR,
                nombre_campo = "EXP_MERMA",
                descrp_campo = "Merma",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FREXTR,
                nombre_campo = "EXP_OBSV",
                descrp_campo = "Observación",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Memo
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FREXTR,
                nombre_campo = "EXP_ESPESOR",
                descrp_campo = "Espesor",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FREXTR,
                nombre_campo = "EXP_UNESP",
                descrp_campo = "Unidad Espesor",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FREXTR,
                nombre_campo = "EXP_CDRTDE",
                descrp_campo = "Cod.Ruta DEST",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FREXTR,
                nombre_campo = "EXP_NMRTDE",
                descrp_campo = "Ruta DEST",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FREXTR,
                nombre_campo = "EXP_VELMAQ",
                descrp_campo = "Velocidad maquina",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Measurement
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FREXTR,
                nombre_campo = "EXP_GOFR",
                descrp_campo = "Gofrado",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 1,
                validValues = new string[] { "Y", "N" },
                validDescription = new string[] { "Si", "No" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FREXTR,
                nombre_campo = "EXP_TPMANGA",
                descrp_campo = "Tipo Manga",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 5,
                validValues = new string[] { "1", "2" },
                validDescription = new string[] { "Abierta 1 lado", "Abierta 2 lados" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FREXTR,
                nombre_campo = "EXP_TERMC",
                descrp_campo = "Termocontraible",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 1,
                validValues = new string[] { "Y", "N" },
                validDescription = new string[] { "Si", "No" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FREXTR,
                nombre_campo = "EXP_OLLA",
                descrp_campo = "Olla (pulgada)",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Measurement
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FREXTR,
                nombre_campo = "EXP_CABEMM",
                descrp_campo = "Cabezal mm",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Measurement
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FREXTR,
                nombre_campo = "EXP_MEDAEX",
                descrp_campo = "Medida a extruir",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Measurement
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FREXTR,
                nombre_campo = "EXP_UMMEDAEX",
                descrp_campo = "UM Med. Extruir",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_None,
                tamano = 10,
                validValues = new string[] { "GR/M2", "CM", "Mils", "µ", "''", "MM" },
                validDescription = new string[] { "Gramos X M2", "Centimetros", "Milesima pulgada", "Micras", "Pulgadas", "Milimetros" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FREXTR,
                nombre_campo = "EXP_AUX1",
                descrp_campo = "AUX1",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FREXTR,
                nombre_campo = "EXP_AUX2",
                descrp_campo = "AUX2",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FREXTR,
                nombre_campo = "EXP_AUX3",
                descrp_campo = "AUX3",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            #endregion
            #region FRIMPR
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_VELMAQ",
                descrp_campo = "Velocidad maquina",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Measurement
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_CODRUT",
                descrp_campo = "Cod.Ruta",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_NOMRUT",
                descrp_campo = "Ruta",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_CODRUTOR",
                descrp_campo = "Cod.Ruta Origen",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_NOMRUTOR",
                descrp_campo = "Ruta Origen",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_RECMAQ",
                descrp_campo = "Maquina Impresión",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 20
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_TPOPRE",
                descrp_campo = "Preparación ",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_MPRIMA",
                descrp_campo = "Materia Prima",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_FORM",
                descrp_campo = "Fórmula",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_SIDES",
                descrp_campo = "Caras",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_BAND",
                descrp_campo = "Bandas",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });

            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_UMFR",
                descrp_campo = "UM Frecuencia",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_None,
                tamano = 10,
                validValues = new string[] { "GR/M2", "CM", "Mils", "µ", "''", "MM"},
                validDescription = new string[] { "Gramos X M2", "Centimetros", "Milesima pulgada", "Micras", "Pulgadas", "Milimetros"}
            });

            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_FREC",
                descrp_campo = "Frecuencia",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });


            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_UMREP",
                descrp_campo = "UM Repeticion",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_None,
                tamano = 10,
                validValues = new string[] { "GR/M2", "CM", "Mils", "µ", "''", "MM" },
                validDescription = new string[] { "Gramos X M2", "Centimetros", "Milesima pulgada", "Micras", "Pulgadas", "Milimetros" }
            });

            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_REPT",
                descrp_campo = "Repeticiones",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });

            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_MTRL",
                descrp_campo = "Material",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 60
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_NROC",
                descrp_campo = "Nro Cilindro mm",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_DESA",
                descrp_campo = "Desarrollo mm",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_NCOL",
                descrp_campo = "Nro Colores",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_COLOR1",
                descrp_campo = "Color 1",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 60
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_COLOR2",
                descrp_campo = "Color 2",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 60
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_COLOR3",
                descrp_campo = "Color 3",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 60
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_COLOR4",
                descrp_campo = "Color 4",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 60
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_COLOR5",
                descrp_campo = "Color 5",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 60
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_COLOR6",
                descrp_campo = "Color 6",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 60
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_COLOR7",
                descrp_campo = "Color 7",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 60
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_COLOR8",
                descrp_campo = "Color 8",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 60
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_MERMA",
                descrp_campo = "Merma",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_OBSV",
                descrp_campo = "Observación",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Memo
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_TIPIM",
                descrp_campo = "Tipo Impresion",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 1,
                validValues = new string[] { "I", "E" },
                validDescription = new string[] { "Interna", "Externa" }

            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_ESPESOR",
                descrp_campo = "Espesor",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_UNESP",
                descrp_campo = "Unidad Espesor",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_SENTIDO",
                descrp_campo = "Sentido Impresion",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 1,
                validValues = new string[] { "1", "2", "3", "4", "5", "6", "7", "8" },
                validDescription = new string[] { "1", "2", "3", "4", "5", "6", "7", "8" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_TITIN1",
                descrp_campo = "Tipo Tinta 1",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_TITIN2",
                descrp_campo = "Tipo Tinta 2",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_TITIN3",
                descrp_campo = "Tipo Tinta 3",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_TITIN4",
                descrp_campo = "Tipo Tinta 4",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_TITIN5",
                descrp_campo = "Tipo Tinta 5",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_TITIN6",
                descrp_campo = "Tipo Tinta 6",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_TITIN7",
                descrp_campo = "Tipo Tinta 7",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_TITIN8",
                descrp_campo = "Tipo Tinta 8",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_UMANC",
                descrp_campo = "UM Ancho",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_None,
                tamano = 10,
                validValues = new string[] { "GR/M2", "CM", "Mils", "µ", "''", "MM" },
                validDescription = new string[] { "Gramos X M2", "Centimetros", "Milesima pulgada", "Micras", "Pulgadas", "Milimetros" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_MEANC",
                descrp_campo = "Med. Ancho",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_UMBAND",
                descrp_campo = "UM Banda",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_None,
                tamano = 10,
                validValues = new string[] { "GR/M2", "CM", "Mils", "µ", "''", "MM" },
                validDescription = new string[] { "Gramos X M2", "Centimetros", "Milesima pulgada", "Micras", "Pulgadas", "Milimetros" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_MEBAND",
                descrp_campo = "Med Banda",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_AUX1",
                descrp_campo = "AUX1",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_AUX2",
                descrp_campo = "AUX2",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRIMPR,
                nombre_campo = "EXP_AUX3",
                descrp_campo = "AUX3",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            #endregion
            #region LAMINA 
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_LAMINA,
                nombre_campo = "EXP_VELMAQ",
                descrp_campo = "Velocidad maquina",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Measurement
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_LAMINA,
                nombre_campo = "EXP_MTRL",
                descrp_campo = "Material",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 60
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_LAMINA,
                nombre_campo = "EXP_CODRUT",
                descrp_campo = "Cod. Ruta",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_LAMINA,
                nombre_campo = "EXP_NOMRUT",
                descrp_campo = "Ruta",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_LAMINA,
                nombre_campo = "EXP_CODRUTOR",
                descrp_campo = "Cod. Ruta Origen",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_LAMINA,
                nombre_campo = "EXP_NOMRUTOR",
                descrp_campo = "Ruta Origen",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_LAMINA,
                nombre_campo = "EXP_RECMAQ",
                descrp_campo = "Maquina Laminado",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_LAMINA,
                nombre_campo = "EXP_TPOPRE",
                descrp_campo = "Preparación ",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_LAMINA,
                nombre_campo = "EXP_MPRIMA",
                descrp_campo = "Materia Prima",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_LAMINA,
                nombre_campo = "EXP_FORM",
                descrp_campo = "Fórmula",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_LAMINA,
                nombre_campo = "EXP_ESPESOR",
                descrp_campo = "Espesor",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_LAMINA,
                nombre_campo = "EXP_UNESP",
                descrp_campo = "Unidad Espesor",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_LAMINA,
                nombre_campo = "EXP_MERMA",
                descrp_campo = "Merma",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_LAMINA,
                nombre_campo = "EXP_OBSV",
                descrp_campo = "Observación",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Memo
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_LAMINA,
                nombre_campo = "EXP_MTRLM2",
                descrp_campo = "M2: Material",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 60
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_LAMINA,
                nombre_campo = "EXP_MPRIM2",
                descrp_campo = "M2: MateriaPrima 2 ",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_LAMINA,
                nombre_campo = "EXP_ESPESM2",
                descrp_campo = "M2: Espesor",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_LAMINA,
                nombre_campo = "EXP_UNESPM2",
                descrp_campo = "M2: Und. Espesor",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_LAMINA,
                nombre_campo = "EXP_MERMAM2",
                descrp_campo = "M2: Merma",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_LAMINA,
                nombre_campo = "EXP_RODILLO",
                descrp_campo = "Rodillo Laminado (mm)",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_LAMINA,
                nombre_campo = "EXP_SENTIDO",
                descrp_campo = "Sentido Laminado",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 1,
                validValues = new string[] { "1", "2", "3", "4", "5", "6", "7", "8" },
                validDescription = new string[] { "1", "2", "3", "4", "5", "6", "7", "8" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_LAMINA,
                nombre_campo = "EXP_AUX1",
                descrp_campo = "AUX1",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_LAMINA,
                nombre_campo = "EXP_AUX2",
                descrp_campo = "AUX2",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_LAMINA,
                nombre_campo = "EXP_AUX3",
                descrp_campo = "AUX3",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            #endregion
            #region FRSELLADO 
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSELA,
                nombre_campo = "EXP_VELMAQ",
                descrp_campo = "Velocidad maquina",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Measurement
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSELA,
                nombre_campo = "EXP_MTRL",
                descrp_campo = "Material",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 60
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSELA,
                nombre_campo = "EXP_CODRUT",
                descrp_campo = "Cod. Ruta",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSELA,
                nombre_campo = "EXP_NOMRUT",
                descrp_campo = "Ruta",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSELA,
                nombre_campo = "EXP_CODRUTOR",
                descrp_campo = "Cod. Ruta Origen",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSELA,
                nombre_campo = "EXP_NOMRUTOR",
                descrp_campo = "Ruta Origen",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSELA,
                nombre_campo = "EXP_RECMAQ",
                descrp_campo = "Maquina Sellado",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSELA,
                nombre_campo = "EXP_TPOPRE",
                descrp_campo = "Preparacion ",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSELA,
                nombre_campo = "EXP_MPRIMA",
                descrp_campo = "Materia Prima",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSELA,
                nombre_campo = "EXP_FORM",
                descrp_campo = "Formula",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSELA,
                nombre_campo = "EXP_ESPESOR",
                descrp_campo = "Espesor",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSELA,
                nombre_campo = "EXP_UNESP",
                descrp_campo = "Unidad Espesor",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSELA,
                nombre_campo = "EXP_MERMA",
                descrp_campo = "Merma",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSELA,
                nombre_campo = "EXP_OBSV",
                descrp_campo = "Observacion",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Memo
            });

            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSELA,
                nombre_campo = "EXP_TROQUEL",
                descrp_campo = "Troquel",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 1,
                validValues = new string[] { "Y", "N" },
                validDescription = new string[] { "Si", "No" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSELA,
                nombre_campo = "EXP_ZIPPER",
                descrp_campo = "Zipper",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 1,
                validValues = new string[] { "Y", "N" },
                validDescription = new string[] { "Si", "No" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSELA,
                nombre_campo = "EXP_TSELLO",
                descrp_campo = "Tipo Sello",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 3,
                validValues = new string[] { "LAT", "FON", "DOS", "FTR", "SET", "SDP", "SEK", "SPO", "STR", "FDB", "CGT", "RMB", "SCO", "COR" },
                validDescription = new string[] { "Lateral", "Fondo", "Doble Sello", "Fondo Troquel", "Sello en T", "Sello en DoyPack", "Sello en K", "Sello Pouch"
                ,"Transversal","Fondo Doble","Corte Guillotina","Sello Rombo","Sello Continuo", "Corte" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSELA,
                nombre_campo = "EXP_TSELL2",
                descrp_campo = "Tipo Sello 2",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 3,
                validValues = new string[] { "LAT", "FON", "DOS", "FTR", "SET", "SDP", "SEK", "SPO", "STR", "FDB", "CGT", "RMB", "SCO", "COR" },
                validDescription = new string[] { "Lateral", "Fondo", "Doble Sello", "Fondo Troquel", "Sello en T", "Sello en DoyPack", "Sello en K", "Sello Pouch"
                ,"Transversal","Fondo Doble","Corte Guillotina","Sello Rombo","Sello Continuo", "Corte" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSELA,
                nombre_campo = "EXP_FUELLE",
                descrp_campo = "Fuelle Fondo",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            //myList.Add(new CampoBean()
            //{
            //    nombre_tabla = TABLE_FRSELA,
            //    nombre_campo = "EXP_SOLAPA",
            //    descrp_campo = "Solapa",
            //    tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
            //    tamano = 50
            //});

            //YA NO ES CHECK
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSELA,
                nombre_campo = "EXP_PESTANA",
                descrp_campo = "Pestaña",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
                //validValues = new string[] { "Y", "N" },
                //validDescription = new string[] { "Si", "No" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSELA,
                nombre_campo = "EXP_CICIBO",
                descrp_campo = "Cinta Cierrabolsa",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 1,
                validValues = new string[] { "Y", "N" },
                validDescription = new string[] { "Si", "No" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSELA,
                nombre_campo = "EXP_MUESCA",
                descrp_campo = "Muesca",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 1,
                validValues = new string[] { "Y", "N" },
                validDescription = new string[] { "Si", "No" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSELA,
                nombre_campo = "EXP_UNFUEL",
                descrp_campo = "Unidad Fuelle",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSELA,
                nombre_campo = "EXP_UNSOLA",
                descrp_campo = "Unidad Solapa",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSELA,
                nombre_campo = "EXP_PRCORTE",
                descrp_campo = "PreCorte",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 1,
                validValues = new string[] { "Y", "N" },
                validDescription = new string[] { "Si", "No" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSELA,
                nombre_campo = "EXP_ULTRASO",
                descrp_campo = "Sello Ultrasonido",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 1,
                validValues = new string[] { "Y", "N" },
                validDescription = new string[] { "Si", "No" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSELA,
                nombre_campo = "EXP_ATROQ",
                descrp_campo = "Asa Troquel",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 1,
                validValues = new string[] { "Y", "N" },
                validDescription = new string[] { "Si", "No" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSELA,
                nombre_campo = "EXP_TROQT",
                descrp_campo = "Troquel en T",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 1,
                validValues = new string[] { "Y", "N" },
                validDescription = new string[] { "Si", "No" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSELA,
                nombre_campo = "EXP_CAVALIER",
                descrp_campo = "Cavalier(mm)",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSELA,
                nombre_campo = "EXP_TRASLAPE",
                descrp_campo = "Traslape (mm)",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSELA,
                nombre_campo = "EXP_WICKET",
                descrp_campo = "Wicket (mm)",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSELA,
                nombre_campo = "EXP_SOLAPA",
                descrp_campo = "Solapa (mm)",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSELA,
                nombre_campo = "EXP_ENVASE",
                descrp_campo = "Envase",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSELA,
                nombre_campo = "EXP_CAIDA",
                descrp_campo = "Caidas",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Numeric,
                tamano = 1,
                validValues = new string[] { "1", "2", "3", "4" },
                validDescription = new string[] { "1", "2", "3", "4" },
                valorPorDef = "1"
            }) ;
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSELA,
                nombre_campo = "EXP_REFUER",
                descrp_campo = "Refuerzo",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Measurement
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSELA,
                nombre_campo = "EXP_ESPREF",
                descrp_campo = "Espesor Refuerzo (mm)",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSELA,
                nombre_campo = "EXP_AUX1",
                descrp_campo = "AUX1",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSELA,
                nombre_campo = "EXP_AUX2",
                descrp_campo = "AUX2",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSELA,
                nombre_campo = "EXP_AUX3",
                descrp_campo = "AUX3",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            #endregion
            #region FRCORT 
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRCORT,
                nombre_campo = "EXP_VELMAQ",
                descrp_campo = "Velocidad maquina",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Measurement
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRCORT,
                nombre_campo = "EXP_CODRUT",
                descrp_campo = "Cod.Ruta",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRCORT,
                nombre_campo = "EXP_NOMRUT",
                descrp_campo = "Ruta",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRCORT,
                nombre_campo = "EXP_CODRUTOR",
                descrp_campo = "Cod.Ruta Origen",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRCORT,
                nombre_campo = "EXP_NOMRUTOR",
                descrp_campo = "Ruta Origen",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRCORT,
                nombre_campo = "EXP_RECMAQ",
                descrp_campo = "Maquina Corte",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRCORT,
                nombre_campo = "EXP_TPOPRE",
                descrp_campo = "Preparación ",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRCORT,
                nombre_campo = "EXP_FORM",
                descrp_campo = "Fórmula",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRCORT,
                nombre_campo = "EXP_MEDCRT",
                descrp_campo = "Medida Corte",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRCORT,
                nombre_campo = "EXP_DIAMT",
                descrp_campo = "Diámetro",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRCORT,
                nombre_campo = "EXP_ANCHO",
                descrp_campo = "Ancho Original",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            /*myList.Add(new CampoBean()
            {
                nombre_tabla = TABLE_FRCORT,
                nombre_campo = "EXP_ANCHMAT",
                descrp_campo = "Ancho Matriz",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });*/                        
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRCORT,
                nombre_campo = "EXP_MERMA",
                descrp_campo = "Merma",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRCORT,
                nombre_campo = "EXP_OBSV",
                descrp_campo = "Observación",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Memo
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRCORT,
                nombre_campo = "EXP_SENTIDO",
                descrp_campo = "Sentido Corte",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 1,
                validValues = new string[] { "1", "2", "3", "4", "5", "6", "7", "8" },
                validDescription = new string[] { "1", "2", "3", "4", "5", "6", "7", "8" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRCORT,
                nombre_campo = "EXP_MPRIMA",
                descrp_campo = "Materia Prima",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRCORT,
                nombre_campo = "EXP_ESPESOR",
                descrp_campo = "Espesor",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRCORT,
                nombre_campo = "EXP_UNESP",
                descrp_campo = "Unidad Espesor",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRCORT,
                nombre_campo = "EXP_PESKG",
                descrp_campo = "Peso Kilos",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRCORT,
                nombre_campo = "EXP_TUCO",
                descrp_campo = "Tuco",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 5,
                validValues = new string[] { "3PLG", "6PLG" },
                validDescription = new string[] { "3 PLG", "6 PLG" }
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRCORT,
                nombre_campo = "EXP_MTRL",
                descrp_campo = "Material",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 60
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRCORT,
                nombre_campo = "EXP_DOBREF",
                descrp_campo = "Doblar y refilar",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 60
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRCORT,
                nombre_campo = "EXP_AUX1",
                descrp_campo = "AUX1",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRCORT,
                nombre_campo = "EXP_AUX2",
                descrp_campo = "AUX2",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRCORT,
                nombre_campo = "EXP_AUX3",
                descrp_campo = "AUX3",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            #endregion
            #region FRHABI 
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRHABI,
                nombre_campo = "EXP_VELMAQ",
                descrp_campo = "Velocidad maquina",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Measurement
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRHABI,
                nombre_campo = "EXP_CODRUT",
                descrp_campo = "Cod.Ruta",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRHABI,
                nombre_campo = "EXP_NOMRUT",
                descrp_campo = "Ruta",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRHABI,
                nombre_campo = "EXP_CODRUTOR",
                descrp_campo = "Cod.Ruta Origen",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRHABI,
                nombre_campo = "EXP_NOMRUTOR",
                descrp_campo = "Ruta Origen",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRHABI,
                nombre_campo = "EXP_RECMAQ",
                descrp_campo = "Maquina Corte",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRHABI,
                nombre_campo = "EXP_TPOPRE",
                descrp_campo = "Preparación ",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRHABI,
                nombre_campo = "EXP_MTRL",
                descrp_campo = "Material",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 60
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRHABI,
                nombre_campo = "EXP_FORM",
                descrp_campo = "Fórmula",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRHABI,
                nombre_campo = "EXP_MEDCRT",
                descrp_campo = "Medida Corte",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRHABI,
                nombre_campo = "EXP_DIAMT",
                descrp_campo = "Diámetro",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRHABI,
                nombre_campo = "EXP_SENTIDO",
                descrp_campo = "Sentido",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRHABI,
                nombre_campo = "EXP_ANCHO",
                descrp_campo = "Ancho Original",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRHABI,
                nombre_campo = "EXP_MERMA",
                descrp_campo = "Merma",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRHABI,
                nombre_campo = "EXP_OBSV",
                descrp_campo = "Observación",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Memo
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRHABI,
                nombre_campo = "EXP_MPRIMA",
                descrp_campo = "Materia Prima",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRHABI,
                nombre_campo = "EXP_ESPESOR",
                descrp_campo = "Espesor",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRHABI,
                nombre_campo = "EXP_UNESP",
                descrp_campo = "Unidad Espesor",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRHABI,
                nombre_campo = "EXP_AUX1",
                descrp_campo = "AUX1",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRHABI,
                nombre_campo = "EXP_AUX2",
                descrp_campo = "AUX2",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRHABI,
                nombre_campo = "EXP_AUX3",
                descrp_campo = "AUX3",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            #endregion
            #region FRREBO 
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRREBO,
                nombre_campo = "EXP_VELMAQ",
                descrp_campo = "Velocidad maquina",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Measurement
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRREBO,
                nombre_campo = "EXP_CODRUT",
                descrp_campo = "Cod.Ruta",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRREBO,
                nombre_campo = "EXP_NOMRUT",
                descrp_campo = "Ruta",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRREBO,
                nombre_campo = "EXP_CODRUTOR",
                descrp_campo = "Cod.Ruta Origen",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRREBO,
                nombre_campo = "EXP_NOMRUTOR",
                descrp_campo = "Ruta Origen",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRREBO,
                nombre_campo = "EXP_RECMAQ",
                descrp_campo = "Maquina Corte",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRREBO,
                nombre_campo = "EXP_TPOPRE",
                descrp_campo = "Preparación ",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRREBO,
                nombre_campo = "EXP_FORM",
                descrp_campo = "Fórmula",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRREBO,
                nombre_campo = "EXP_MEDCRT",
                descrp_campo = "Medida Corte",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRREBO,
                nombre_campo = "EXP_DIAMT",
                descrp_campo = "Diámetro",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRREBO,
                nombre_campo = "EXP_SENTIDO",
                descrp_campo = "Sentido",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRREBO,
                nombre_campo = "EXP_MTRL",
                descrp_campo = "Material",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 60
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRREBO,
                nombre_campo = "EXP_ANCHO",
                descrp_campo = "Ancho Original",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRREBO,
                nombre_campo = "EXP_OBSV",
                descrp_campo = "Observación",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Memo
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRREBO,
                nombre_campo = "EXP_MPRIMA",
                descrp_campo = "Materia Prima",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRREBO,
                nombre_campo = "EXP_ESPESOR",
                descrp_campo = "Espesor",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRREBO,
                nombre_campo = "EXP_UNESP",
                descrp_campo = "Unidad Espesor",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRREBO,
                nombre_campo = "EXP_MERMA",
                descrp_campo = "Merma",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo = SAPbobsCOM.BoFldSubTypes.st_Quantity
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRREBO,
                nombre_campo = "EXP_AUX1",
                descrp_campo = "AUX1",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRREBO,
                nombre_campo = "EXP_AUX2",
                descrp_campo = "AUX2",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRREBO,
                nombre_campo = "EXP_AUX3",
                descrp_campo = "AUX3",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            #endregion
            #region FRSERV 
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSERV,
                nombre_campo = "EXP_CODRUT",
                descrp_campo = "Cod.Ruta",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSERV,
                nombre_campo = "EXP_NOMRUT",
                descrp_campo = "Ruta",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSERV,
                nombre_campo = "EXP_CODRUTOR",
                descrp_campo = "Cod.Ruta Origen",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSERV,
                nombre_campo = "EXP_NOMRUTOR",
                descrp_campo = "Ruta Origen",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSERV,
                nombre_campo = "EXP_SERV",
                descrp_campo = "Servicio",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 50
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSERV,
                nombre_campo = "EXP_IMPORTE",
                descrp_campo = "Importe",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Float,
                subtipo_campo=SAPbobsCOM.BoFldSubTypes.st_Price
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSERV,
                nombre_campo = "EXP_OBSV",
                descrp_campo = "Observación",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Memo
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSERV,
                nombre_campo = "EXP_AUX1",
                descrp_campo = "AUX1",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSERV,
                nombre_campo = "EXP_AUX2",
                descrp_campo = "AUX2",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
            });
            myList.Add(new CampoModel()
            {
                nombre_tabla = TABLE_FRSERV,
                nombre_campo = "EXP_AUX3",
                descrp_campo = "AUX3",
                tipo_campo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                tamano = 100
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
                name = "RUTA_LISTAMATERIALES",
                tableName = TABLE_CABE,
                canCancel = SAPbobsCOM.BoYesNoEnum.tNO,
                canClose = SAPbobsCOM.BoYesNoEnum.tYES,
                canDelete = SAPbobsCOM.BoYesNoEnum.tYES,
                childTables = new string[] { TABLE_FRRUTA, TABLE_FORMUL, TABLE_SUBPRD, TABLE_INDUCT, TABLE_FREXTR, TABLE_FRIMPR, TABLE_LAMINA,  TABLE_FRCORT, TABLE_FRHABI, TABLE_FRREBO, TABLE_FRSELA, TABLE_FRSERV },
                canCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO,
                canFind = SAPbobsCOM.BoYesNoEnum.tYES,
                canLog = SAPbobsCOM.BoYesNoEnum.tNO,
                objectType = SAPbobsCOM.BoUDOObjType.boud_MasterData,
                enableEnhancedForm = SAPbobsCOM.BoYesNoEnum.tNO,
                rebuildEnhancedForm = SAPbobsCOM.BoYesNoEnum.tNO
            };
            return myObj;
        }

        #endregion


    }
}
