using System;
using System.Globalization;
using System.Threading;
using System.Windows.Forms;

namespace Gastos
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // Configurar cultura para usar $ como símbolo de moneda
            var culture = new CultureInfo("es-AR");
            culture.NumberFormat.CurrencySymbol = "$";
            Thread.CurrentThread.CurrentCulture = culture;
            Thread.CurrentThread.CurrentUICulture = culture;
            
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new FormMenuPrincipal());
        }
    }
}
