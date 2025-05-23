using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using AddonProduccionEnsDes.commons;
using AddonProduccionEnsDes.conexion;

namespace AddonProduccionEnsDes.entity
{
    public class Monedas : FormCommon
    {
        public string codigo { get; set; }
        public string descripcion { get; set; }

        public List<Monedas> listar()
        {
            var list = new List<Monedas>();
            SAPbobsCOM.Recordset oRS = null;
            try
            {
                oRS = Conexion.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRS.DoQuery("select \"CurrCode\", \"CurrName\" from \"OCRN\" ");

                if (oRS.RecordCount > 0)
                {
                    while (!oRS.EoF)
                    {
                        list.Add(new Monedas
                        {
                            codigo = oRS.Fields.Item("CurrCode").Value.ToString().Trim(),
                            descripcion = oRS.Fields.Item("CurrName").Value.ToString().Trim()
                        });

                        oRS.MoveNext();
                    }
                }
                else
                    StatusMessageInfo("No se encontraron indicadores de impuesto.");
            }
            catch (Exception ex)
            {
                StatusMessageError("Monedas > listar() > " + ex.Message);
            }
            finally
            {
                LiberarObjetoGenerico(oRS);
            }

            return list;
        }
    }
}
