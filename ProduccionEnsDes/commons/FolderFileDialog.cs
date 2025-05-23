using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AddonProduccionEnsDes.commons
{
    public class FolderFileDialog
    {
        public string ruta;

        public string FindFile()
        {
            try
            {
                var variable_temp = string.Empty;
                var explorer = new FileExplorer();
                explorer.ShowFolderBrowser(true, "Archivo anexo |*.xlsx");

                if (explorer.Error)
                {
                    //ShowMessage(explorer.LastException.Message);
                }
                else
                {
                    if (explorer.Files.Length > 0)
                    {
                        variable_temp = explorer.Files[0];
                    }
                }

                if (!string.IsNullOrEmpty(variable_temp))
                {
                    return variable_temp;
                }

            }
            catch (Exception ex)
            {
                //logger.Error(ex.Message, ex);
            }

            return string.Empty;
        }
    }
}
