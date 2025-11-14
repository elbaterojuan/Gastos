using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Gastos.Models;
using Microsoft.Office.Interop.Excel;

namespace Gastos.Services
{
    /// <summary>
    /// Servicio para manejar operaciones con Excel de forma asíncrona
    /// </summary>
    public class ExcelService : IDisposable
    {
        private Microsoft.Office.Interop.Excel.Application _excelApp;
        private Workbooks _workbooks;
        private Workbook _workbook;
        private readonly string _carpeta;
        private readonly string _archivo;

        public ExcelService(string carpeta, string archivo)
        {
            _carpeta = carpeta;
            _archivo = archivo;
            InicializarExcel();
        }

        private void InicializarExcel()
        {
            try
            {
                // Intentar obtener instancia existente
                _excelApp = (Microsoft.Office.Interop.Excel.Application)Marshal.GetActiveObject("Excel.Application");
            }
            catch
            {
                // Crear nueva instancia
                _excelApp = new Microsoft.Office.Interop.Excel.Application();
            }

            _workbooks = _excelApp.Workbooks;
            _excelApp.WindowState = XlWindowState.xlMaximized;
            _excelApp.Visible = true;

            AbrirWorkbook();
        }

        private void AbrirWorkbook()
        {
            try
            {
                _workbook = _workbooks.get_Item(_archivo);
            }
            catch
            {
                string rutaCompleta = Path.Combine(_carpeta, _archivo);
                _workbook = _workbooks.Open(rutaCompleta);
            }
        }

        /// <summary>
        /// Agrega un gasto al archivo de Excel
        /// </summary>
        public async Task<bool> AgregarGastoAsync(Gasto gasto)
        {
            return await Task.Run(() =>
            {
                try
                {
                    Worksheet hoja = ObtenerHoja(gasto.Fecha);
                    int fila = ObtenerUltimaFila(hoja);

                    ((Range)hoja.Cells[fila, 1]).Value = gasto.Fecha.ToString("M/d/yyyy");
                    ((Range)hoja.Cells[fila, 2]).Value = gasto.Categoria;
                    ((Range)hoja.Cells[fila, 3]).Value = gasto.Monto;
                    ((Range)hoja.Cells[fila, 4]).Value = gasto.QuienPago;
                    ((Range)hoja.Cells[fila, 5]).Value = gasto.EsProporcional ? "SI" : "NO";
                    ((Range)hoja.Cells[fila, 8]).Value = gasto.Cuotas == 1 
                        ? gasto.Comentarios 
                        : $"{gasto.Comentarios} {gasto.Cuotas}";

                    return true;
                }
                catch (Exception ex)
                {
                    throw new Exception($"Error al agregar gasto: {ex.Message}", ex);
                }
            });
        }

        /// <summary>
        /// Obtiene todos los gastos de un mes específico
        /// </summary>
        public async Task<List<Gasto>> ObtenerGastosMesAsync(DateTime fecha)
        {
            return await Task.Run(() =>
            {
                var gastos = new List<Gasto>();
                try
                {
                    Worksheet hoja = ObtenerHoja(fecha);
                    int ultimaFila = hoja.UsedRange.Rows.Count;

                    for (int i = 2; i <= ultimaFila; i++) // Asumiendo que fila 1 es header
                    {
                        Range celda1 = (Range)hoja.Cells[i, 1];
                        if (celda1.Value == null) continue;

                        var gasto = new Gasto
                        {
                            Fecha = DateTime.Parse(celda1.Value.ToString()),
                            Categoria = ((Range)hoja.Cells[i, 2]).Value?.ToString() ?? "",
                            Monto = decimal.Parse(((Range)hoja.Cells[i, 3]).Value?.ToString() ?? "0"),
                            QuienPago = ((Range)hoja.Cells[i, 4]).Value?.ToString() ?? "",
                            EsProporcional = ((Range)hoja.Cells[i, 5]).Value?.ToString() == "SI",
                            Comentarios = ((Range)hoja.Cells[i, 8]).Value?.ToString() ?? ""
                        };

                        gastos.Add(gasto);
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception($"Error al obtener gastos: {ex.Message}", ex);
                }

                return gastos;
            });
        }

        /// <summary>
        /// Calcula el resumen de gastos para un mes
        /// </summary>
        public async Task<ResumenMensual> ObtenerResumenMensualAsync(DateTime fecha)
        {
            var gastos = await ObtenerGastosMesAsync(fecha);
            var resumen = new ResumenMensual
            {
                Mes = fecha.Month,
                Año = fecha.Year,
                CantidadGastos = gastos.Count,
                TotalGastos = gastos.Sum(g => g.Monto),
                PromedioGastos = gastos.Any() ? gastos.Average(g => g.Monto) : 0
            };

            // Gastos por categoría
            resumen.GastosPorCategoria = gastos
                .GroupBy(g => g.Categoria)
                .ToDictionary(g => g.Key, g => g.Sum(x => x.Monto));

            // Gastos por persona
            resumen.GastosPorPersona = gastos
                .GroupBy(g => g.QuienPago)
                .ToDictionary(g => g.Key, g => g.Sum(x => x.Monto));

            return resumen;
        }

        private Worksheet ObtenerHoja(DateTime fecha)
        {
            string nombreMes = System.Globalization.CultureInfo.GetCultureInfo("es-ES")
                .DateTimeFormat.GetMonthName(fecha.Month);
            nombreMes = char.ToUpper(nombreMes[0]) + nombreMes.Substring(1);
            string nombreHoja = $"{nombreMes}-{fecha:yy}";

            return (Worksheet)_workbook.Worksheets[nombreHoja];
        }

        private int ObtenerUltimaFila(Worksheet hoja)
        {
            for (int i = 1; i <= hoja.UsedRange.Rows.Count; i++)
            {
                Range cell1 = (Range)hoja.Cells[i, 1];
                Range cell2 = (Range)hoja.Cells[i, 2];
                Range cell3 = (Range)hoja.Cells[i, 3];

                if (cell1.Value == null && cell2.Value == null && cell3.Value == null)
                {
                    return i;
                }
            }

            return hoja.UsedRange.Rows.Count + 1;
        }

        public void Dispose()
        {
            if (_excelApp != null)
            {
                try
                {
                    _workbooks?.Close();
                    _excelApp.Quit();

                    if (_workbooks != null) Marshal.ReleaseComObject(_workbooks);
                    if (_excelApp != null) Marshal.ReleaseComObject(_excelApp);
                }
                catch { }
            }
        }
    }
}
