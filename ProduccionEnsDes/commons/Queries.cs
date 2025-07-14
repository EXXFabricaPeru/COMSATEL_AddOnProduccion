using System;
using System.Text;

namespace AddonProduccionEnsDes.commons
{
    public class Queries
    {
        #region _Attributes_

        private static StringBuilder m_sSQL = new StringBuilder();

        #endregion

        #region _Functions_
        public static string GetRecurrentePendiente(string fecha)
        {
            m_sSQL.Length = 0;
            m_sSQL.Append("SELECT R1.\"DocEntry\",R1.\"LineId\", R0.\"U_EXC_ORDREF\",R0.\"U_EXC_ORDBAS\" ");
            m_sSQL.Append("FROM \"@EXC_RECUR\" R0 ");
            m_sSQL.Append("JOIN \"@EXC_RECUR1\" R1 ON R0.\"DocEntry\"=R1.\"DocEntry\" ");
            m_sSQL.AppendFormat("WHERE R1.\"U_EXC_ESTADO\" = 'P' AND R1.\"U_EXC_FECINI\"<= '{0}' AND R0.\"U_EXC_ORDBAS\"!=-1 ", fecha);
            return m_sSQL.ToString();
        }
        public static string GetDetalleRecurrente(string DocEntry)
        {
            m_sSQL.Length = 0;
            m_sSQL.Append("Select DISTINCT T1.\"U_EXC_ORDREC\" \"Instancia\",U1.\"Descr\" \"Estado\",T1.\"U_EXC_FECINI\" \"Fecha\",T1.\"U_EXC_TOTDOC\" \"Total\" ");
            m_sSQL.Append("from \"@EXC_RECUR1\" \"T1\" ");
            m_sSQL.Append("JOIN \"@EXC_RECUR\" \"T0\" ON T1.\"DocEntry\"= T0.\"DocEntry\" ");
            m_sSQL.Append("JOIN \"UFD1\" \"U1\" ON  U1.\"TableID\" = '@EXC_RECUR1' AND T1.\"U_EXC_ESTADO\"= U1.\"FldValue\" ");
            m_sSQL.AppendFormat("WHERE T0.\"DocEntry\"='{0}' ", DocEntry);
            m_sSQL.Append("ORDER BY T1.\"U_EXC_FECINI\" ");
            return m_sSQL.ToString();
        }

        public static string GetRecurrentes(string DocEntry)
        {
            m_sSQL.Length = 0;
            m_sSQL.Append("Select  DISTINCT \"DocEntry\",\"U_EXC_IDENT\" \"ID\",\"U_EXC_FECINI\" \"Inicio\",\"U_EXC_FECFIN\" \"Fin\",\"Name\" \"Per\",\"U_EXC_DIAREP\" \"Dia\",\"U_EXC_ORDBAS\" \"Base\"  ");
            m_sSQL.Append("from \"@EXC_RECUR\" ");
            m_sSQL.Append("JOIN \"@EXC_PERIFAC\" ON  \"U_EXC_PERIFAC\"=\"Code\" ");
            m_sSQL.AppendFormat("WHERE \"U_EXC_CONTID\"='{0}' ", DocEntry);
            return m_sSQL.ToString();
        }
        public static string GetDetalleProductoEnsamble(string ItemCode)
        {
            m_sSQL.Length = 0;
            m_sSQL.Append("SELECT T1.\"Code\", T1.\"ItemName\"  ");
            m_sSQL.Append("FROM OITT T0 ");
            m_sSQL.Append("JOIN ITT1 T1 ON T0.\"Code\"=T1.\"Father\" ");
            m_sSQL.Append("JOIN OITM IT ON T1.\"Code\"= IT.\"ItemCode\" ");
            m_sSQL.AppendFormat("WHERE T0.\"Code\"='{0}' AND IT.\"U_EXC_EQPPROD\" = 'Y';", ItemCode);

            string query = m_sSQL.ToString();
            return query;
        }
        public static string GetStatusOrder(string DocEntry)
        {
            m_sSQL.Length = 0;
            m_sSQL.AppendFormat("SELECT \"Status\" FROM OWOR WHERE \"DocEntry\" = {0}; ", DocEntry);

            string query = m_sSQL.ToString();
            return query;
        }

        public static string GetAbsEntrySerie(string ItemCode, string IMEI)
        {
            m_sSQL.Length = 0;
            m_sSQL.Append("SELECT \"AbsEntry\" ");
            m_sSQL.Append("FROM OSRN ");
            m_sSQL.AppendFormat("WHERE \"ItemCode\" = '{0}' AND \"U_EXC_IMEI\" = '{1}';", ItemCode, IMEI.Replace("'", "''"));

            string query = m_sSQL.ToString();
            return query;
        }

        public static string GetInternalNumberSerie(string ItemCode, string IMEI)
        {
            m_sSQL.Length = 0;
            m_sSQL.Append("SELECT \"DistNumber\" ");
            m_sSQL.Append("FROM OSRN ");
            m_sSQL.AppendFormat("WHERE \"ItemCode\" = '{0}' AND \"U_EXC_IMEI\" = '{1}';", ItemCode, IMEI.Replace("'", "''"));

            string query = m_sSQL.ToString();
            return query;
        }
        public static string GetDetalleProductoProducir(string ItemCode)
        {
            m_sSQL.Length = 0;
            m_sSQL.Append("SELECT T1.\"Code\", T1.\"ItemName\"  ");
            m_sSQL.Append("FROM OITT T0 ");
            m_sSQL.Append("JOIN ITT1 T1 ON T0.\"Code\"=T1.\"Father\" ");
            m_sSQL.Append("JOIN OITM IT ON T1.\"Code\"= IT.\"ItemCode\" ");
            m_sSQL.AppendFormat("WHERE T0.\"Code\"='{0}'", ItemCode);

            string query = m_sSQL.ToString();
            return query;
        }
        public static string GetDetalleItemEnsamble(string ItemCode, string Serie)//OPTIMIZAR
        {
            m_sSQL.Length = 0;
            m_sSQL.Append("SELECT TOP 1 \"U_EXC_MARCA\",\"U_EXC_MODELO\",\"U_EXC_IMEI\",\"U_EXC_OPERAD\"  ");
            m_sSQL.Append("FROM OSRN SE ");
            m_sSQL.AppendFormat("WHERE SE.\"ItemCode\"='{0}' AND SE.\"DistNumber\"='{1}' ", ItemCode, Serie);
            return m_sSQL.ToString();
        }

        public static string GetAlmacenOT(string DocEntry)
        {
            m_sSQL.Length = 0;
            m_sSQL.Append("SELECT T0.\"Warehouse\"  ");
            m_sSQL.Append("FROM OWOR T0 ");
            m_sSQL.AppendFormat("WHERE T0.\"DocEntry\"={0};", DocEntry);

            string query = m_sSQL.ToString();
            return query;
        }

        public static string GetDetalleDesensamble(string ItemCode, string Serie)
        {
            m_sSQL.Length = 0;
            m_sSQL.Append("SELECT ");
            //m_sSQL.Append("EN.\"U_EXC_CEQP\", ");
            //m_sSQL.Append("EN.\"U_EXC_DEQP\", ");
            m_sSQL.Append("SP.\"DistNumber\" \"U_EXC_SEQP\", ");
            m_sSQL.Append("SP.\"U_EXC_MARCA\",  ");
            m_sSQL.Append("SP.\"U_EXC_MODELO\",  ");
            //m_sSQL.Append("SP.\"U_EXC_CCHI\",  ");
            // m_sSQL.Append("SP.\"U_EXC_DCHI\",  ");
            //m_sSQL.Append("SP.\"DistNumber\" \"U_EXC_SCHI\",  ");
            m_sSQL.Append("SP.\"U_EXC_IMEI\",  ");
            m_sSQL.Append("SP.\"U_EXC_OPERAD\",  ");
            m_sSQL.Append("'008-ALM' \"U_EXC_ALMI\",  ");
            //m_sSQL.Append("E0.\"U_EXC_ALMA\" \"U_EXC_ALMS\",  ");

            m_sSQL.Append("SP.\"U_EXC_FIRMW\",  ");
            m_sSQL.Append("SP.\"U_EXC_APN\",  ");
            m_sSQL.Append("SP.\"U_EXC_TIPIP\",  ");
            m_sSQL.Append("SP.\"U_EXC_IP\",  ");
            m_sSQL.Append("SP.\"U_EXC_LINEA\",  ");
            m_sSQL.Append("SP.\"U_EXC_FOTA\",  ");
            m_sSQL.Append("SP.\"U_EXC_PAQDATOS\",  ");
            m_sSQL.Append("SP.\"U_EXC_SIMCARD\",  ");
            m_sSQL.Append("SP.\"U_EXC_PRODPOR\",  ");
            //Nuevo 20220201
            m_sSQL.Append("SP.\"U_EXC_LINTEL\",  ");
            m_sSQL.Append("SP.\"U_EXC_PROTC\",  ");
            m_sSQL.Append("SP.\"U_EXC_IP\",  ");

            //Nuevo 20241222
            m_sSQL.Append("SP.\"U_EXC_CACC1\",  ");
            m_sSQL.Append("(SELECT \"ItemName\" FROM OITM A WHERE A.\"ItemCode\" = SP.\"U_EXC_CACC1\") AS \"U_EXC_DACC1\",  ");
            m_sSQL.Append("SP.\"U_EXC_CACC2\",  ");
            m_sSQL.Append("(SELECT \"ItemName\" FROM OITM A WHERE A.\"ItemCode\" = SP.\"U_EXC_CACC2\") AS \"U_EXC_DACC2\"  ");
            m_sSQL.Append("FROM OSRN SP  ");
            //m_sSQL.Append("JOIN \"@EXC_ENS1\" EN ON EN.\"U_EXC_NSER\" =SP.\"DistNumber\" ");
            //m_sSQL.Append("JOIN \"@EXC_ENSA\" E0 ON EN.\"DocEntry\" =E0.\"DocEntry\" ");
            //m_sSQL.Append("LEFT JOIN OSRN SC ON  SC.\"ItemCode\"= EN.\"U_EXC_CCHI\" AND SC.\"SysNumber\"=EN.\"U_EXC_SCHI\" ");
            //m_sSQL.Append("LEFT JOIN OSRN SE ON  SE.\"ItemCode\"= EN.\"U_EXC_CEQP\" AND SE.\"SysNumber\"=EN.\"U_EXC_SEQP\" ");
            m_sSQL.AppendFormat("WHERE SP.\"ItemCode\"='{0}' AND SP.\"DistNumber\"='{1}' ", ItemCode, Serie);
            string query = m_sSQL.ToString();
            return query;
        }

        public static string GetWhsSerie(string ItemCode, string Serie)//OPTIMIZAR
        {
            m_sSQL.Length = 0;
            m_sSQL.Append("SELECT TOP 1  \"WhsCode\" \"Value\"  ");
            m_sSQL.Append("FROM OSRQ SE ");
            m_sSQL.AppendFormat("WHERE SE.\"ItemCode\"='{0}' AND SE.\"SysNumber\"='{1}' AND SE.\"Quantity\">0", ItemCode, Serie);
            return m_sSQL.ToString();
        }
        public static string GetSerieTerminado(string Marca, string Modelo)//OPTIMIZAR
        {
            m_sSQL.Length = 0;
            m_sSQL.Append("SELECT \"U_EXC_PREFIJO\" || LPAD (\"U_EXC_CORR\" + 1,5,'0') \"Value\"  ");
            m_sSQL.Append("FROM \"@EXC_SERIEPROD\" SE ");
            m_sSQL.AppendFormat("WHERE SE.\"U_EXC_MARCA\"='{0}' AND SE.\"U_EXC_MODELO\"='{1}' ", Marca, Modelo);
            return m_sSQL.ToString();
        }

        public static string UpdateSerieUDO(string DocEntry, string OT, string Serie)//OPTIMIZAR
        {
            m_sSQL.Length = 0;
            m_sSQL.AppendFormat("UPDATE \"@EXC_ENS1\" SET \"U_EXC_NSER\"='{0}' ", Serie);
            m_sSQL.AppendFormat("WHERE \"DocEntry\"='{0}' AND \"U_EXC_ORDT\"='{1}' ", DocEntry, OT);
            return m_sSQL.ToString();
        }

        public static string UpdateSerieEquipo(string ItemCode, string Serie, SAPbouiCOM.DBDataSource dsDETA, int i)//OPTIMIZAR
        {
            m_sSQL.Length = 0;
            m_sSQL.AppendFormat("UPDATE OSRN ");
            m_sSQL.AppendFormat("SET \"U_EXC_IMEI\"='{0}', \"U_EXC_MARCA\"='{1}', \"U_EXC_MODELO\"='{2}', \"U_EXC_PRODPOR\"='{3}', \"U_EXC_FIRMW\"='{4}' "
                                , dsDETA.GetValue("U_EXC_IMEI", i)
                                , dsDETA.GetValue("U_EXC_MARC", i)
                                , dsDETA.GetValue("U_EXC_MODE", i)
                                , dsDETA.GetValue("U_EXC_PROP", i)
                                , dsDETA.GetValue("U_EXC_FWAR", i));
            m_sSQL.AppendFormat("WHERE \"ItemCode\" = '{0}' AND \"DistNumber\" = '{1}' ", ItemCode, Serie);
            return m_sSQL.ToString();
        }

        public static string UpdateSerieChip(string ItemCode, string Serie, SAPbouiCOM.DBDataSource dsDETA, int i)//OPTIMIZAR
        {
            m_sSQL.Length = 0;
            m_sSQL.Append("UPDATE OSRN ");
            m_sSQL.AppendFormat("SET \"U_EXC_SIMCARD\"='{0}', \"U_EXC_LINTEL\"='{1}', \"U_EXC_OPERAD\"='{2}', \"U_EXC_PAQDATOS\"='{3}', \"U_EXC_APN\"='{4}',\"U_EXC_TIPIP\"='{5}' "
                                , dsDETA.GetValue("U_EXC_SIMC", i)
                                , dsDETA.GetValue("U_EXC_LTEL", i)
                                , dsDETA.GetValue("U_EXC_OPER", i)
                                , dsDETA.GetValue("U_EXC_PQDA", i)
                                , dsDETA.GetValue("U_EXC_DAPN", i)
                                , dsDETA.GetValue("U_EXC_TIIP", i));
            m_sSQL.AppendFormat("WHERE \"ItemCode\" = '{0}' AND \"DistNumber\" = '{1}' ", ItemCode, Serie);
            return m_sSQL.ToString();
        }

        public static string UpdateSerieTerminado(string Marca, string Modelo)//OPTIMIZAR
        {
            m_sSQL.Length = 0;
            m_sSQL.Append("UPDATE \"@EXC_SERIEPROD\" SET \"U_EXC_CORR\"=\"U_EXC_CORR\"+1 ");
            m_sSQL.AppendFormat("WHERE \"U_EXC_MARCA\"='{0}' AND \"U_EXC_MODELO\"='{1}' ", Marca, Modelo);
            return m_sSQL.ToString();
        }

        public static string CheckPreviousSeries(string ItemCode, string Serie, bool isChip)//corrgiendo
        {
            m_sSQL.Length = 0;
            m_sSQL.Append("SELECT COUNT(*) \"Value\"  ");
            m_sSQL.Append("FROM \"@EXC_ENS1\" E1 ");
            m_sSQL.Append("JOIN \"@EXC_ENSA\" E0 ON E1.\"DocEntry\" = E0.\"DocEntry\" ");
            //m_sSQL.Append("AND E0.\"U_EXC_ESTA\" !='F' ");
            m_sSQL.Append("AND IFNULL(NULLIF(E1.\"U_EXC_NSER\",''),'') = '' ");
            m_sSQL.AppendFormat("WHERE E1.\"{2}\"='{0}' AND E1.\"{3}\"='{1}' ", ItemCode, Serie, (isChip ? "U_EXC_CCHI" : "U_EXC_CEQP"), (isChip ? "U_EXC_SCHI" : "U_EXC_SEQP"));
            m_sSQL.Append("AND E0.\"Status\"='O' ");

            string query = m_sSQL.ToString();

            return query;
        }
        public static string CheckPreviousSeriesProd(string ItemCode, string Serie)//OPTIMIZAR
        {
            //m_sSQL.Length = 0;
            //m_sSQL.Append("SELECT COUNT(*) \"Value\"  ");
            //m_sSQL.Append("FROM \"@EXC_ENS1\" D1");
            //m_sSQL.AppendFormat("WHERE \"U_EXC_CPRO\"='{0}' AND \"U_EXC_NSER\"='{1}' ", ItemCode, Serie);
            //return m_sSQL.ToString();

            m_sSQL.Length = 0;
            m_sSQL.Append("SELECT COUNT(*) \"Value\"  ");
            m_sSQL.Append("FROM \"@EXC_DES1\" D1 ");
            m_sSQL.Append("JOIN \"@EXC_DESE\" D0 ON D1.\"DocEntry\" = D0.\"DocEntry\" ");
            m_sSQL.AppendFormat("WHERE D1.\"U_EXC_CPRO\"='{0}' AND \"U_EXC_NSER\"='{1}' ", ItemCode, Serie);
            m_sSQL.Append("AND D0.\"Status\"='O' ");
            string query = m_sSQL.ToString();
            return query;
        }
        public static string GetSerieItem(string ItemCode = "")
        {
            m_sSQL.Length = 0;
            m_sSQL.Append("SELECT DISTINCT  SE.\"SysNumber\" \"Value\",SE.\"DistNumber\" \"Name\"  ");
            m_sSQL.Append("FROM OSRN SE ");

            if (!string.IsNullOrEmpty(ItemCode))
            {
                m_sSQL.AppendFormat("WHERE SE.\"ItemCode\"='{0}' ", ItemCode);
                m_sSQL.Append("AND NULLIF(SE.\"U_EXC_IMEI\",'') IS NOT NULL AND ((((NULLIF(SE.\"U_EXC_OPERAD\",'') IS NOT NULL ) ) OR ( (NULLIF(SE.\"U_EXC_MARCA\",'') IS NOT NULL AND NULLIF(SE.\"U_EXC_MODELO\",'') IS NOT NULL ) ) ))");
            }
            m_sSQL.Append("ORDER BY SE.\"DistNumber\" ");

            string query = m_sSQL.ToString();
            return query;
        }

        public static string GetSerieItemDS(string ItemCode = "")
        {
            m_sSQL.Length = 0;
            m_sSQL.Append("SELECT DISTINCT  SE.\"SysNumber\" \"Value\",SE.\"DistNumber\" \"Name\"  ");
            m_sSQL.Append("FROM OSRN SE ");

            if (!string.IsNullOrEmpty(ItemCode))
            {
                m_sSQL.AppendFormat("WHERE SE.\"ItemCode\"='{0}' ", ItemCode);
                m_sSQL.Append("AND NULLIF(SE.\"U_EXC_IMEI\",'') IS NOT NULL AND ((((NULLIF(SE.\"U_EXC_OPERAD\",'') IS NOT NULL ) ) OR ( (NULLIF(SE.\"U_EXC_MARCA\",'') IS NOT NULL AND NULLIF(SE.\"U_EXC_MODELO\",'') IS NOT NULL ) ) ))");
            }
            m_sSQL.Append("ORDER BY SE.\"DistNumber\" ");

            string query = m_sSQL.ToString();
            return query;
        }

        public static string GetSeriesChip(string ItemCode)
        {
            m_sSQL.Length = 0;
            m_sSQL.Append("select TOP 1 T1.\"Code\", T1.\"ItemName\"  ");
            m_sSQL.Append("FROM OITT T0 ");
            m_sSQL.Append("JOIN ITT1 T1 ON T0.\"Code\"=T1.\"Father\" ");
            m_sSQL.AppendFormat("JOIN OITM IT ON T1.\"Code\"= IT.\"ItemCode\" AND IT.\"ItmsGrpCod\"!='{0}' ", "253");
            m_sSQL.AppendFormat("WHERE T0.\"Code\"='{0}' ", ItemCode);
            return m_sSQL.ToString();
        }
        public static string GetDetalleLlamada(string DocEntry)
        {
            m_sSQL.Length = 0;
            m_sSQL.Append("SELECT \"U_EXC_CODART\" \"Articulo\", \"U_EXC_DESCRIP\" \"Descripcion\", \"U_EXC_CANTIDAD\" \"Cantidad\", \"U_EXC_ESTADO\" \"Estado\", \"U_EXC_IDDOC\" \"Doc.\", \"U_EXC_OBJTYPE\" \"Tipo\", \"U_EXC_FECHDOC\" \"Fecha\", \"U_EXC_NUMDOC\" \"DocNum\"  ");
            m_sSQL.Append("from \"@EXC_DETLLAM\" ");
            m_sSQL.AppendFormat("WHERE \"U_EXC_CALLID\"='{0}' ", DocEntry);
            return m_sSQL.ToString();
        }
        public static string Querieseries()
        {
            m_sSQL.Length = 0;
            //m_sSQL.AppendFormat("SELECT \"Series\", \"SeriesName\", \"DocSubType\" FROM NNM1 where \"ObjectCode\" IN ('{0}')  AND \"Locked\" = 'N' and \"DocSubType\" IN ('--','IB')", 13);
            m_sSQL.Append("SELECT \"Series\", \"SeriesName\", \"DocSubType\"");
            m_sSQL.AppendFormat(" FROM NNM1 where \"ObjectCode\" IN ('{0}')  AND \"Locked\" = 'N' AND \"Indicator\" = '{1}' and \"DocSubType\" IN ('--','IB')", 13, DateTime.Now.Year);
            return m_sSQL.ToString();
        }

        public static string GetPagoProrrata(string entry, string table, string taxCodePro)
        {
            m_sSQL.Length = 0;
            m_sSQL.Append("Select P0.\"DocEntry\",P0.\"DocDate\", P0.\"NumAtCard\", P0.\"CardCode\", P1.\"AcctCode\", P0.\"DocType\", ");
            m_sSQL.Append(" P1.\"OcrCode\",P1.\"OcrCode2\",P1.\"OcrCode3\", ");
            m_sSQL.Append(" P0.\"DocCur\", ");
            m_sSQL.AppendFormat(" IFNULL(NULLIF(P1.\"U_EXD_PRORRATA\",0),CASE WHEN P0.\"DocCur\"='{0}' THEN P0.\"TotalExpns\" ELSE P0.\"TotalExpFC\" END) \"Prorrota\" ", Constants.MAIN_CURR);
            m_sSQL.AppendFormat(" FROM {0} \"P0\"", table);
            m_sSQL.AppendFormat(" JOIN {0} \"P1\" ON P0.\"DocEntry\"=P1.\"DocEntry\"", table.Replace("O", "") + "1");
            m_sSQL.AppendFormat(" WHERE P0.\"DocEntry\"='{0}' AND P1.\"TaxCode\"='{1}'", entry, taxCodePro);

            return m_sSQL.ToString();
        }

        public static string GetCheckCFGAux(string entry)
        {
            m_sSQL.Length = 0;
            m_sSQL.AppendFormat("SELECT TOP 1 IFNULL(\"{0}\",'N') \"Check\" FROM \"@{1}\"", entry, data_schema.SCConfig.TABLE_CABE);
            return m_sSQL.ToString();
        }

        public static string GetNextFormCode(string itemcode)
        {
            m_sSQL.Length = 0;
            m_sSQL.AppendFormat("select top 1 * from(SELECT TOP 1  LPAD(SUBSTRING(\"Code\", length('{0}')+1)+1,2,'0') \"Value\" FROM \"@EXP_OFRM\" where \"Code\" like '{0}%' ", itemcode);
            m_sSQL.AppendFormat("union all ");
            m_sSQL.AppendFormat("SELECT '00' \"Value\" FROM dummy) order by \"Value\" desc; ");
            return m_sSQL.ToString();
        }

        public static string GetCFGValue(string entry)
        {
            m_sSQL.Length = 0;
            m_sSQL.AppendFormat("SELECT TOP 1 IFNULL(\"{0}\",'') \"Value\" FROM \"@{1}\"", entry, data_schema.SCConfig.TABLE_CABE);
            return m_sSQL.ToString();

        }
        public static string GetUDFValue(string table, string campo, string cod)
        {
            m_sSQL.Length = 0;
            m_sSQL.AppendFormat("SELECT T1.\"Descr\" \"Value\" ");
            m_sSQL.AppendFormat("FROM CUFD \"T0\" ");
            m_sSQL.AppendFormat("JOIN UFD1 \"T1\" ON T0.\"TableID\"=T1.\"TableID\" AND T0.\"FieldID\"=T1.\"FieldID\" ");
            m_sSQL.AppendFormat("WHERE T0.\"TableID\" = '@{0}' AND T0.\"AliasID\" = '{1}' AND T1.\"FldValue\" = '{2}'", table.Trim(), campo.Trim(), cod.Trim());
            return m_sSQL.ToString();
        }

        public static string GetFormulacion(string itemcode)
        {
            m_sSQL.Length = 0;
            m_sSQL.AppendFormat("SELECT ");
            m_sSQL.AppendFormat("FO0.\"U_EXP_QTT\" \"Qty\" ");
            m_sSQL.AppendFormat("FROM OITM \"ITM\" ");
            m_sSQL.AppendFormat("JOIN \"@{0}\" \"FO0\" ON ITM.\"U_EXP_FORM\"=FO0.\"Code\" ", data_schema.SCFormulacion.TABLE_CABE);
            m_sSQL.AppendFormat("WHERE ITM.\"ItemCode\"='{0}'", itemcode);

            return m_sSQL.ToString();
        }
        public static string GetFormulacionDetalle(string itemcode)
        {
            m_sSQL.Length = 0;
            m_sSQL.AppendFormat("CALL \"EXP_FormulacionListaMateriales\" ('{0}')", itemcode);

            return m_sSQL.ToString();
        }

        public static string GetComboTax()
        {
            m_sSQL.Length = 0;
            m_sSQL.AppendFormat("SELECT  \"Code\" \"Value\",\"Name\" \"Name\" FROM \"OSTC\" WHERE \"Lock\"='N'");
            return m_sSQL.ToString();
        }

        public static string GetComboRutas()
        {
            m_sSQL.Length = 0;
            m_sSQL.AppendFormat("SELECT  \"Code\" \"Value\",\"Desc\" \"Name\" FROM \"ORST\"");
            return m_sSQL.ToString();
        }

        public static string GetComboMateriales()
        {
            m_sSQL.Length = 0;
            m_sSQL.AppendFormat("SELECT  \"Name\" \"Value\",\"Name\" \"Name\" FROM \"@EXP_TIPMAT\"");
            return m_sSQL.ToString();
        }
        public static string GetComboTinta()
        {
            m_sSQL.Length = 0;
            m_sSQL.AppendFormat("SELECT  \"Code\" \"Value\",\"Name\" \"Name\" FROM \"@EXP_TINTA\"");
            return m_sSQL.ToString();
        }

        public static string GetComboUnidadMedidas()
        {
            m_sSQL.Length = 0;
            m_sSQL.AppendFormat("SELECT  \"Code\" \"Value\",\"Name\" \"Name\" FROM \"@EXP_CONMED\"");
            return m_sSQL.ToString();
        }

        public static string GetRefile(string ItemCode)
        {
            m_sSQL.Length = 0;
            m_sSQL.Append("SELECT TOP 1 IFNULL(IFNULL(NULLIF(\"U_EXP_REFANC\",0),NULLIF(\"U_EXP_REFLAR\",0)),0) \"Value\" ");
            m_sSQL.Append("FROM OITM ");
            m_sSQL.AppendFormat("WHERE \"ItemCode\" = '{0}'", ItemCode);
            return m_sSQL.ToString();
        }


        public static string GetRutaValues(string cod)
        {
            m_sSQL.Length = 0;
            m_sSQL.Append("SELECT RU.\"Desc\" \"Value\", IFNULL(RU.\"U_EXP_SUBPRD\",'') \"SUBPRD\", IFNULL(RU.\"U_EXP_SCRAP\",'') \"SCRAP\", IFNULL(RU.\"U_EXP_REFILE\",'') \"REFILE\" ");
            m_sSQL.Append(", IFNULL(IT.\"ItemCode\",'') \"INDUCTOR\" ");
            m_sSQL.Append("FROM \"ORST\" \"RU\" ");
            m_sSQL.Append("LEFT JOIN \"OITM\" \"IT\" ON RU.\"Code\"=IT.\"U_EXP_RIND\" ");
            m_sSQL.AppendFormat("WHERE RU.\"Code\"='{0}'", cod);
            return m_sSQL.ToString();
        }

        public static string GetComboPeriodos()
        {
            m_sSQL.Length = 0;
            m_sSQL.AppendFormat("SELECT TOP 6 distinct \"LineId\",\"U_Fecha\" \"Value\",\"U_Fecha\" \"Name\" FROM \"@EXD_FACT_PROR\" ORDER BY  \"LineId\" DESC");
            return m_sSQL.ToString();
        }


        public static string ConsultaTablaConfiguracion(SAPbobsCOM.BoDataServerTypes bo_ServerTypes, string NAddon, string Version, bool Ordenamiento)
        {
            m_sSQL.Length = 0;

            switch (bo_ServerTypes)
            {
                case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                    m_sSQL.AppendFormat("SELECT * FROM \"@{0}\"", NAddon.ToUpper());
                    if (NAddon != "" || Version != "")
                    {
                        m_sSQL.Append(" WHERE ");
                        if (NAddon != "")
                        {
                            m_sSQL.AppendFormat("\"Name\" Like '{0}%'", NAddon);
                            if (Version != "") m_sSQL.AppendFormat(" AND \"Code\" = '{0}'", Version);
                        }
                        else if (Version != "") m_sSQL.AppendFormat("\"Code\" = '{0}'", Version);
                    }
                    if (Ordenamiento) m_sSQL.Append(" ORDER BY LENGTH(\"Code\") DESC, \"Code\" DESC");

                    break;
                default:
                    m_sSQL.AppendFormat("SELECT * FROM [@{0}]", NAddon.ToUpper());
                    if (NAddon != "" || Version != "")
                    {
                        m_sSQL.Append(" WHERE ");
                        if (NAddon != "")
                        {
                            m_sSQL.AppendFormat("Name Like '{0}%'", NAddon);
                            if (Version != "") m_sSQL.AppendFormat(" AND Code = '{0}'", Version);
                        }
                        else if (Version != "") m_sSQL.AppendFormat("Code = '{0}'", Version);
                    }
                    if (Ordenamiento) m_sSQL.Append(" ORDER BY LEN(Code) DESC, Code DESC");
                    break;
            }

            return m_sSQL.ToString();
        }

        #endregion

    }
}
