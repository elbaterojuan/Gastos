using System;

namespace Gastos.Models
{
    /// <summary>
    /// Representa un gasto individual registrado en el sistema
    /// </summary>
    public class Gasto
    {
        public DateTime Fecha { get; set; }
        public string Categoria { get; set; }
        public decimal Monto { get; set; }
        public string QuienPago { get; set; }
        public bool EsProporcional { get; set; }
        public string Comentarios { get; set; }

        public Gasto()
        {
            Fecha = DateTime.Now;
            EsProporcional = false;
        }

        public override string ToString()
        {
            return $"{Fecha:dd/MM/yyyy} - {Categoria} - ${Monto:N2}";
        }
    }
}
