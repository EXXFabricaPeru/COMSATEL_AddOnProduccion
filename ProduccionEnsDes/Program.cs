using System;
using System.Windows.Forms;
using AddonProduccionEnsDes.commons;
using AddonProduccionEnsDes.conexion;

namespace AddonProduccionEnsDes
{
    static class Program
    {
        /// <summary>
        /// Punto de entrada principal para la aplicación.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Conexion oConexion = new Conexion();

            if ((oConexion != null) && (Conexion.company.Connected))
            {
                //DataStructure oEstructuraDatos = new DataStructure();
                //GC.KeepAlive(oConexion);
                FormCommon.StatusMessageSuccess("Inicio satisfactorio."); 
                Application.Run();
            }
            else
                Application.Exit();

            Application.Run();
        }
    }
}
