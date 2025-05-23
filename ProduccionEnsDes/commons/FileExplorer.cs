using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace AddonProduccionEnsDes.commons
{
    public class FileExplorer
    {
        public string[] Files;
        public string Files2;
        public bool Error;
        public Exception LastException;


        public void ShowFolderBrowser(bool multiple, string filter = null)
        {
            try
            {
                Thread ShowFolderBrowserThread = new System.Threading.Thread(() => ShowFolderBrowserX(filter, multiple));

                if (ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Unstarted)
                {
                    ShowFolderBrowserThread.SetApartmentState(ApartmentState.STA);
                    ShowFolderBrowserThread.Start();
                    ShowFolderBrowserThread.Join();
                    try
                    {
                        ShowFolderBrowserThread.Abort();
                    }
                    catch (ThreadAbortException ex)
                    {

                    }
                }
                else if (ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Stopped)
                {
                    ShowFolderBrowserThread.Start();
                    ShowFolderBrowserThread.Join();
                }
            }
            catch (Exception ex)
            {
                ex.ToString();
            }
        }

        private void ShowFolderBrowserX(string Filter, bool multiple)
        {
            try
            {
                Form nws = new Form();
                OpenFileDialog MyTest = new OpenFileDialog();

                MyTest.Multiselect = multiple;

                if (!string.IsNullOrWhiteSpace(Filter))
                    MyTest.Filter = Filter;

                nws.Size = new System.Drawing.Size(1, 1);
                nws.Show();
                nws.BringToFront();
                nws.Focus();
                if (MyTest.ShowDialog(nws) == DialogResult.OK)
                {
                    Files = MyTest.FileNames;
                }
                else
                {
                    Files = new string[] { };
                }
                nws.Dispose();
                Error = false;
            }
            catch (Exception ex)
            {
                Error = true;
                LastException = ex;
            }
        }
    }
}
