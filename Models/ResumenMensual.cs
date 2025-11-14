using System;
using System.Collections.Generic;

namespace Gastos.Models
{
    /// <summary>
    /// Resumen estadístico de gastos por mes
    /// </summary>
    public class ResumenMensual
    {
        public int Mes { get; set; }
        public int Año { get; set; }
        public decimal TotalGastos { get; set; }
        public decimal PromedioGastos { get; set; }
        public int CantidadGastos { get; set; }
        public Dictionary<string, decimal> GastosPorCategoria { get; set; }
        public Dictionary<string, decimal> GastosPorPersona { get; set; }

        public ResumenMensual()
        {
            GastosPorCategoria = new Dictionary<string, decimal>();
            GastosPorPersona = new Dictionary<string, decimal>();
        }

        public string NombreMes
        {
            get
            {
                return new DateTime(Año, Mes, 1).ToString("MMMM yyyy");
            }
        }
    }
}
