using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AddonProduccionEnsDes.commons
{
    public class WindowWrapper : System.Windows.Forms.IWin32Window
    {
        private IntPtr _hwnd;
        // Constructor
        public WindowWrapper(IntPtr handle)
        {
            _hwnd = handle;
        }
        // Property
        public virtual IntPtr Handle
        {
            get { return _hwnd; }
        }


    }
}
