using AddonProduccionEnsDes.conexion;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddonProduccionEnsDes.commons
{
    public class LayoutService
    {
        public static void CrearLayout(string typeName, string addonName, string formName, string menuUID, string rptPath)
        {
            string typeCode = string.Empty;

            try
            {
                typeCode = ObtenerTypeCode(typeName, addonName, formName);
                      
                if (typeCode == string.Empty)
                {
                    typeCode = CrearTipo(typeName, addonName, formName, menuUID);

                    if (typeCode != string.Empty)
                        AsignarLayout(typeCode, formName, rptPath);
                }
                else
                {
                    if (!Asignado(typeCode, formName))
                        AsignarLayout(typeCode, formName, rptPath);
                }
            }
            catch (Exception e)
            {
                throw new Exception("No se pudo asignar el reporte. Error : " + e.Message);
            }
        }

        
        private static bool Asignado(string typeCode, string formName)
        {
            Recordset recordset = null;

            try
            {
                string query = $"select \"DocCode\" from \"RDOC\" WHERE \"TypeCode\" = '{typeCode}' and \"DocName\" = '{formName}'";
                recordset = Conexion.company.GetBusinessObject(BoObjectTypes.BoRecordset);
                recordset.DoQuery(query);

                return recordset.RecordCount > 0;
            }
            catch (Exception)
            {
                throw;
            }
            finally { LiberarObjetoGenerico(recordset); };

        }

        private static void AsignarLayout(string typeCode, string formName, string rptPath)
        {
            try
            {
                ReportLayoutsService oLayoutservice = (ReportLayoutsService)Conexion.company.GetCompanyService().GetBusinessService(ServiceTypes.ReportLayoutsService);
                ReportLayout oReport = (ReportLayout)oLayoutservice.GetDataInterface(ReportLayoutsServiceDataInterfaces.rlsdiReportLayout);

                oReport.Name = formName;
                oReport.TypeCode = typeCode;
                oReport.Author = Conexion.company.UserName;
                oReport.Category = ReportLayoutCategoryEnum.rlcCrystal;

                string newReportCode;
                ReportLayoutParams oNewReportParams = oLayoutservice.AddReportLayout(oReport);
                newReportCode = oNewReportParams.LayoutCode;

                CompanyService oCompanyService = Conexion.company.GetCompanyService();
                BlobParams oBlobParams = (BlobParams)oCompanyService.GetDataInterface(CompanyServiceDataInterfaces.csdiBlobParams);

                oBlobParams.Table = "RDOC";
                oBlobParams.Field = "Template";
                BlobTableKeySegment oKeySegment = oBlobParams.BlobTableKeySegments.Add();
                oKeySegment.Name = "DocCode";
                oKeySegment.Value = newReportCode;

                Blob oBlob = (Blob)oCompanyService.GetDataInterface(CompanyServiceDataInterfaces.csdiBlob);

                FileStream oFile = new FileStream(rptPath, FileMode.Open);
                int fileSize = (int)oFile.Length;
                byte[] buf = new byte[fileSize];

                oFile.Read(buf, 0, fileSize);
                oFile.Close();
                oBlob.Content = Convert.ToBase64String(buf, 0, fileSize);

                oCompanyService.SetBlob(oBlobParams, oBlob);

                ReportTypesService rptTypeService = (ReportTypesService)Conexion.company.GetCompanyService().GetBusinessService(ServiceTypes.ReportTypesService);
                ReportTypeParams prms = rptTypeService.GetDataInterface(ReportTypesServiceDataInterfaces.rtsReportTypeParams);
                prms.TypeCode = typeCode;
                ReportType type = rptTypeService.GetReportType(prms);
                type.DefaultReportLayout = newReportCode;
                rptTypeService.UpdateReportType(type);
            }
            catch (Exception)
            {
                throw;
            }
        }

        private static string CrearTipo(string typeName, string addonName, string formName, string menuUID)
        {
            string typeCode = string.Empty;

            try
            {
                ReportTypesService rptTypeService = (ReportTypesService)Conexion.company.GetCompanyService().GetBusinessService(ServiceTypes.ReportTypesService);
                ReportType newType = (ReportType)rptTypeService.GetDataInterface(ReportTypesServiceDataInterfaces.rtsReportType);
                newType.TypeName = typeName;
                newType.AddonName = addonName;
                newType.AddonFormType = formName;
                newType.MenuID = menuUID;
                ReportTypeParams newTypeParam = rptTypeService.AddReportType(newType);

                typeCode = newTypeParam.TypeCode;
            }
            catch (Exception)
            {
            }

            return typeCode;
        }

        public static string ObtenerTypeCode(string typeName, string addonName, string formName)
        {
            Recordset recordset = null;

            try
            {
                string query = $"select \"CODE\" from \"RTYP\" where \"NAME\" = '{typeName}' and \"ADD_NAME\" = '{addonName}' and \"FRM_TYPE\" = '{formName}'";
                recordset = Conexion.company.GetBusinessObject(BoObjectTypes.BoRecordset);
                recordset.DoQuery(query);

                if (!recordset.EoF)
                    return recordset.Fields.Item(0).Value.ToString();

                return string.Empty;
            }
            catch (Exception)
            {
                throw;
            }
            finally { LiberarObjetoGenerico(recordset); }
        }

        internal static void LiberarObjetoGenerico(Object objeto)
        {
            try
            {
                if (objeto != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objeto);

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(Constants.PREFIX_MSG_ADDON + " Error Liberando Objeto: " + ex.Message);
            }
        }
    }
}
