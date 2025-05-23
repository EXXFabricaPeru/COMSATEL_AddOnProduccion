using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using AddonProduccionEnsDes.commons;
using AddonProduccionEnsDes.conexion;

namespace AddonProduccionEnsDes.entity
{
    public class Igv : FormCommon
    {
        public string codigo { get; set; }
        public string descripcion { get; set; }

        public List<Igv> listar()
        {
            var list = new List<Igv>();
            SAPbobsCOM.Recordset oRS = null;
            try
            {
                oRS = Conexion.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRS.DoQuery("select \"Code\", \"Name\" from \"OSTC\" where \"Lock\" = 'N'");

                if (oRS.RecordCount > 0)
                {
                    while (!oRS.EoF)
                    {
                        list.Add(new Igv
                        {
                            codigo = oRS.Fields.Item("Code").Value.ToString().Trim(),
                            descripcion = oRS.Fields.Item("Name").Value.ToString().Trim()
                        });

                        oRS.MoveNext();
                    }
                }
                else
                    StatusMessageInfo("No se encontraron indicadores de impuesto.");
            }
            catch (Exception ex)
            {
                StatusMessageError("Igv > listar() > " + ex.Message);
            }
            finally
            {
                LiberarObjetoGenerico(oRS);
            }

            return list;
        }
    }
}
