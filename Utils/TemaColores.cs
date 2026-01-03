using System.Drawing;

namespace Gastos.Utils
{
    /// <summary>
    /// Paleta de colores moderna para la aplicación
    /// </summary>
    public static class TemaColores
    {
        // Colores principales
        public static readonly Color PrimarioAzul = ColorTranslator.FromHtml("#2563EB");
        public static readonly Color PrimarioAzulOscuro = ColorTranslator.FromHtml("#1E40AF");
        public static readonly Color SecundarioVerde = ColorTranslator.FromHtml("#10B981");
        public static readonly Color AccentoNaranja = ColorTranslator.FromHtml("#F59E0B");

        // Fondos
        public static readonly Color FondoClaro = ColorTranslator.FromHtml("#F9FAFB");
        public static readonly Color FondoBlanco = Color.White;
        public static readonly Color FondoGris = ColorTranslator.FromHtml("#F3F4F6");

        // Texto
        public static readonly Color TextoOscuro = ColorTranslator.FromHtml("#111827");
        public static readonly Color TextoGris = ColorTranslator.FromHtml("#6B7280");
        public static readonly Color TextoClaro = ColorTranslator.FromHtml("#9CA3AF");

        // Estados
        public static readonly Color Exito = ColorTranslator.FromHtml("#10B981");
        public static readonly Color Error = ColorTranslator.FromHtml("#EF4444");
        public static readonly Color Advertencia = ColorTranslator.FromHtml("#F59E0B");
        public static readonly Color Info = ColorTranslator.FromHtml("#3B82F6");

        // Bordes
        public static readonly Color BordeGris = ColorTranslator.FromHtml("#E5E7EB");
        public static readonly Color BordeAzul = ColorTranslator.FromHtml("#BFDBFE");
    }
}
