using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Gastos.Models;
using OfficeOpenXml;

namespace Gastos.Services
{
    /// <summary>
    /// Servicio para manejar operaciones con Excel usando EPPlus (sin necesidad de tener Excel instalado)
    /// </summary>
    public class ExcelService : IDisposable
    {
        private readonly string _excelPath;

        public ExcelService(string carpeta, string archivo)
        {
            // Configurar licencia EPPlus (uso no comercial)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
            _excelPath = Path.Combine(carpeta, archivo);
            
            // Crear carpeta si no existe
            if (!Directory.Exists(carpeta))
            {
                Directory.CreateDirectory(carpeta);
            }
            
            // Crear archivo si no existe
            if (!File.Exists(_excelPath))
            {
                CrearArchivoExcel();
            }
        }

        private void CrearArchivoExcel()
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Gastos");
                
                // Crear encabezados
                worksheet.Cells[1, 1].Value = "Fecha";
                worksheet.Cells[1, 2].Value = "Categoría";
                worksheet.Cells[1, 3].Value = "Monto";
                worksheet.Cells[1, 4].Value = "Quién pagó?";
                worksheet.Cells[1, 5].Value = "Gasto Proporcional?";
                // Columnas 6 y 7 ocultas (no las creamos)
                worksheet.Cells[1, 8].Value = "Comentarios";
                
                // Aplicar estilo a encabezados
                using (var range = worksheet.Cells[1, 1, 1, 8])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(37, 99, 235));
                    range.Style.Font.Color.SetColor(System.Drawing.Color.White);
                }
                
                // Ajustar ancho de columnas
                worksheet.Column(1).Width = 12;
                worksheet.Column(2).Width = 20;
                worksheet.Column(3).Width = 12;
                worksheet.Column(4).Width = 15;
                worksheet.Column(5).Width = 18;
                worksheet.Column(8).Width = 35;
                
                package.SaveAs(new FileInfo(_excelPath));
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
                    using (var package = new ExcelPackage(new FileInfo(_excelPath)))
                    {
                        // Obtener el nombre de la hoja según el mes del gasto
                        string nombreHoja = ObtenerNombreHoja(gasto.Fecha);
                        var worksheet = package.Workbook.Worksheets[nombreHoja];
                        
                        if (worksheet == null)
                        {
                            // Crear todas las hojas faltantes desde la última existente hasta la requerida
                            CrearHojasFaltantes(package, gasto.Fecha);
                            
                            // Ahora obtener la hoja que acabamos de crear
                            worksheet = package.Workbook.Worksheets[nombreHoja];
                        }
                        
                        // Encontrar última fila
                        int lastRow = worksheet.Dimension?.End.Row ?? 1;
                        int newRow = lastRow + 1;
                        
                        // Agregar datos
                        worksheet.Cells[newRow, 1].Value = gasto.Fecha;
                        worksheet.Cells[newRow, 1].Style.Numberformat.Format = "dd/mm/yyyy";
                        worksheet.Cells[newRow, 2].Value = gasto.Categoria;
                        worksheet.Cells[newRow, 3].Value = gasto.Monto;
                        worksheet.Cells[newRow, 4].Value = gasto.QuienPago;
                        worksheet.Cells[newRow, 5].Value = gasto.EsProporcional ? "Sí" : "No";
                        // Columnas 6 y 7 ocultas (las dejamos vacías)
                        worksheet.Cells[newRow, 8].Value = gasto.Comentarios;
                        
                        // Formatear monto como moneda
                        worksheet.Cells[newRow, 3].Style.Numberformat.Format = "$#,##0.00";
                        
                        // Forzar recálculo de fórmulas
                        package.Workbook.Calculate();
                        // Marcar para recalcular al abrir en Excel
                        package.Workbook.CalcMode = ExcelCalcMode.Automatic;
                        package.Workbook.Properties.Modified = DateTime.Now;
                        
                        package.Save();
                        return true;
                    }
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
                    if (!File.Exists(_excelPath))
                        return gastos;
                    
                    using (var package = new ExcelPackage(new FileInfo(_excelPath)))
                    {
                        // Buscar la hoja del mes de diferentes formas
                        var worksheet = BuscarHojaMes(package, fecha);
                        
                        if (worksheet == null)
                        {
                            // Hoja no encontrada para este mes
                            System.Diagnostics.Debug.WriteLine($"No se encontró hoja para {fecha:MMMM yyyy}");
                            return gastos;
                        }
                        
                        if (worksheet.Dimension == null)
                        {
                            // La hoja está vacía
                            return gastos;
                        }
                        
                        int rowCount = worksheet.Dimension.End.Row;
                        
                        System.Diagnostics.Debug.WriteLine($"Leyendo hoja: {worksheet.Name} - Filas: {rowCount}");
                        
                        for (int row = 2; row <= rowCount; row++)
                        {
                            try
                            {
                                var fechaCell = worksheet.Cells[row, 1].Value;
                                if (fechaCell == null)
                                    continue;
                                
                                DateTime fechaGasto;
                                
                                // Intentar parsear la fecha de diferentes formas
                                if (fechaCell is DateTime)
                                {
                                    fechaGasto = (DateTime)fechaCell;
                                }
                                else if (fechaCell is double)
                                {
                                    // Excel almacena fechas como números
                                    fechaGasto = DateTime.FromOADate((double)fechaCell);
                                }
                                else
                                {
                                    var fechaStr = fechaCell.ToString();
                                    if (!DateTime.TryParse(fechaStr, out fechaGasto))
                                        continue;
                                }
                                
                                var montoCell = worksheet.Cells[row, 3].Value;
                                decimal monto = 0;
                                
                                if (montoCell != null)
                                {
                                    if (montoCell is double || montoCell is int)
                                    {
                                        monto = Convert.ToDecimal(montoCell);
                                    }
                                    else
                                    {
                                        decimal.TryParse(montoCell.ToString(), out monto);
                                    }
                                }
                                
                                var gasto = new Gasto
                                {
                                    Fecha = fechaGasto,
                                    Categoria = worksheet.Cells[row, 2].Value?.ToString() ?? "",
                                    Monto = monto,
                                    QuienPago = worksheet.Cells[row, 4].Value?.ToString() ?? "",
                                    EsProporcional = worksheet.Cells[row, 5].Value?.ToString()?.ToLower() == "sí" || 
                                             worksheet.Cells[row, 5].Value?.ToString()?.ToLower() == "si" ||
                                             worksheet.Cells[row, 5].Value?.ToString()?.ToLower() == "true" ||
                                             worksheet.Cells[row, 5].Value?.ToString() == "1",
                                    Comentarios = worksheet.Cells[row, 8].Value?.ToString() ?? ""
                                };
                                
                                gastos.Add(gasto);
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"Error en fila {row}: {ex.Message}");
                                continue;
                            }
                        }
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
        /// Busca la hoja del mes en el Excel, probando diferentes formatos
        /// </summary>
        private ExcelWorksheet BuscarHojaMes(ExcelPackage package, DateTime fecha)
        {
            string[] meses = { "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
                              "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre" };
            
            string nombreMes = meses[fecha.Month - 1];
            string año = fecha.ToString("yy");
            
            // Probar diferentes formatos de nombre
            string[] formatosPosibles = new[]
            {
                $"{nombreMes}-{año}",           // Mayo-25
                $"{nombreMes} {año}",           // Mayo 25
                $"{nombreMes}-20{año}",         // Mayo-2025
                $"{nombreMes} 20{año}",         // Mayo 2025
                nombreMes                        // Mayo
            };
            
            foreach (var formato in formatosPosibles)
            {
                var worksheet = package.Workbook.Worksheets[formato];
                if (worksheet != null)
                {
                    System.Diagnostics.Debug.WriteLine($"Hoja encontrada: {formato}");
                    return worksheet;
                }
            }
            
            // Si no se encuentra, listar todas las hojas disponibles para debug
            System.Diagnostics.Debug.WriteLine($"Hojas disponibles:");
            foreach (var ws in package.Workbook.Worksheets)
            {
                System.Diagnostics.Debug.WriteLine($"  - {ws.Name}");
            }
            
            return null;
        }

        /// <summary>
        /// Obtiene el nombre de la hoja según el formato "Mes-AA"
        /// </summary>
        private string ObtenerNombreHoja(DateTime fecha)
        {
            string[] meses = { "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
                              "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre" };
            
            string nombreMes = meses[fecha.Month - 1];
            string año = fecha.ToString("yy");
            
            return $"{nombreMes}-{año}";
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
                CantidadGastos = gastos.Count
            };

            // Leer valores calculados por Excel desde la fila 2
            await LeerTotalesDesdeExcelAsync(fecha, resumen);

            // Gastos por categoría
            resumen.GastosPorCategoria = gastos
                .GroupBy(g => g.Categoria)
                .ToDictionary(g => g.Key, g => g.Sum(x => x.Monto));

            // Leer deudas desde las columnas M y N del Excel
            await LeerDeudasAsync(fecha, resumen);

            // Leer sueldos desde las columnas J y K del Excel
            await LeerSueldosAsync(fecha, resumen);

            return resumen;
        }

        /// <summary>
        /// Lee los totales calculados por Excel desde la fila 2 (J2, K2, L2)
        /// </summary>
        private async Task LeerTotalesDesdeExcelAsync(DateTime fecha, ResumenMensual resumen)
        {
            await Task.Run(() =>
            {
                try
                {
                    if (!File.Exists(_excelPath))
                        return;

                    using (var package = new ExcelPackage(new FileInfo(_excelPath)))
                    {
                        // Calcular todas las fórmulas antes de leer
                        package.Workbook.Calculate();
                        
                        var worksheet = BuscarHojaMes(package, fecha);
                        if (worksheet == null || worksheet.Dimension == null)
                            return;

                        // J2 = Total Gastos
                        var totalGastos = worksheet.Cells[2, 10].Value;
                        if (totalGastos != null)
                        {
                            if (totalGastos is double || totalGastos is int)
                                resumen.TotalGastado = Convert.ToDecimal(totalGastos);
                            else if (decimal.TryParse(totalGastos.ToString(), out decimal total))
                                resumen.TotalGastado = total;
                        }

                        // Calcular promedio (si hay gastos)
                        if (resumen.CantidadGastos > 0)
                            resumen.PromedioGasto = resumen.TotalGastado / resumen.CantidadGastos;

                        // K2 = Total Andrea
                        var totalAndrea = worksheet.Cells[2, 11].Value;
                        if (totalAndrea != null)
                        {
                            decimal valorAndrea = 0;
                            if (totalAndrea is double || totalAndrea is int)
                                valorAndrea = Convert.ToDecimal(totalAndrea);
                            else if (decimal.TryParse(totalAndrea.ToString(), out decimal ta))
                                valorAndrea = ta;

                            if (!resumen.GastosPorPersona.ContainsKey("Andrea"))
                                resumen.GastosPorPersona["Andrea"] = valorAndrea;
                        }

                        // L2 = Total Juan
                        var totalJuan = worksheet.Cells[2, 12].Value;
                        if (totalJuan != null)
                        {
                            decimal valorJuan = 0;
                            if (totalJuan is double || totalJuan is int)
                                valorJuan = Convert.ToDecimal(totalJuan);
                            else if (decimal.TryParse(totalJuan.ToString(), out decimal tj))
                                valorJuan = tj;

                            if (!resumen.GastosPorPersona.ContainsKey("Juan"))
                                resumen.GastosPorPersona["Juan"] = valorJuan;
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error al leer totales desde Excel: {ex.Message}");
                }
            });
        }

        /// <summary>
        /// Lee las deudas desde las columnas M y N del Excel
        /// </summary>
        private async Task LeerDeudasAsync(DateTime fecha, ResumenMensual resumen)
        {
            await Task.Run(() =>
            {
                try
                {
                    if (!File.Exists(_excelPath))
                        return;

                    using (var package = new ExcelPackage(new FileInfo(_excelPath)))
                    {
                        // Calcular todas las fórmulas antes de leer
                        package.Workbook.Calculate();
                        
                        var worksheet = BuscarHojaMes(package, fecha);
                        if (worksheet == null || worksheet.Dimension == null)
                            return;

                        // Las deudas suelen estar en la primera fila de datos (fila 2)
                        // Columna M (13) = "Debe Andrea"
                        // Columna N (14) = "Debe Juan"
                        var debeAndrea = worksheet.Cells[2, 13].Value;
                        var debeJuan = worksheet.Cells[2, 14].Value;

                        decimal montoAndrea = 0;
                        decimal montoJuan = 0;

                        if (debeAndrea != null && decimal.TryParse(debeAndrea.ToString(), out montoAndrea))
                        {
                            if (montoAndrea > 0)
                            {
                                resumen.DeudorNombre = "Andrea";
                                resumen.DeudorMonto = montoAndrea;
                                return;
                            }
                        }

                        if (debeJuan != null && decimal.TryParse(debeJuan.ToString(), out montoJuan))
                        {
                            if (montoJuan > 0)
                            {
                                resumen.DeudorNombre = "Juan";
                                resumen.DeudorMonto = montoJuan;
                                return;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error al leer deudas: {ex.Message}");
                }
            });
        }

        /// <summary>
        /// Lee los sueldos desde las columnas J y K del Excel (filas 7-11)
        /// </summary>
        private async Task LeerSueldosAsync(DateTime fecha, ResumenMensual resumen)
        {
            await Task.Run(() =>
            {
                try
                {
                    if (!File.Exists(_excelPath))
                        return;

                    using (var package = new ExcelPackage(new FileInfo(_excelPath)))
                    {
                        // Calcular todas las fórmulas antes de leer
                        package.Workbook.Calculate();
                        
                        var worksheet = BuscarHojaMes(package, fecha);
                        if (worksheet == null || worksheet.Dimension == null)
                            return;

                        // Columna K (11) - Todos los valores están en la columna K
                        // Fila 7 = Sueldo Andrea (K7)
                        // Fila 8 = Sueldo Juan (K8)
                        // Fila 9 = Total (K9)
                        // Fila 10 = Porcentaje Andrea (K10)
                        // Fila 11 = Porcentaje Juan (K11)

                        // Leer todos los valores de la columna K
                        var sueldoAndrea = worksheet.Cells[7, 11].Value;  // K7
                        var sueldoJuan = worksheet.Cells[8, 11].Value;    // K8
                        var total = worksheet.Cells[9, 11].Value;         // K9
                        var pctAndrea = worksheet.Cells[10, 11].Value;    // K10
                        var pctJuan = worksheet.Cells[11, 11].Value;      // K11

                        if (sueldoAndrea != null)
                        {
                            if (sueldoAndrea is double || sueldoAndrea is int)
                                resumen.SueldoAndrea = Convert.ToDecimal(sueldoAndrea);
                            else if (decimal.TryParse(sueldoAndrea.ToString(), out decimal sa))
                                resumen.SueldoAndrea = sa;
                        }

                        if (sueldoJuan != null)
                        {
                            if (sueldoJuan is double || sueldoJuan is int)
                                resumen.SueldoJuan = Convert.ToDecimal(sueldoJuan);
                            else if (decimal.TryParse(sueldoJuan.ToString(), out decimal sj))
                                resumen.SueldoJuan = sj;
                        }

                        if (total != null)
                        {
                            if (total is double || total is int)
                                resumen.SueldoTotal = Convert.ToDecimal(total);
                            else if (decimal.TryParse(total.ToString(), out decimal t))
                                resumen.SueldoTotal = t;
                        }

                        if (pctAndrea != null)
                        {
                            decimal valor = 0;
                            if (pctAndrea is double || pctAndrea is int)
                                valor = Convert.ToDecimal(pctAndrea);
                            else if (decimal.TryParse(pctAndrea.ToString(), out decimal pa))
                                valor = pa;
                            
                            // Multiplicar por 100 para obtener el porcentaje entero
                            resumen.PorcentajeAndrea = valor * 100;
                        }

                        if (pctJuan != null)
                        {
                            decimal valor = 0;
                            if (pctJuan is double || pctJuan is int)
                                valor = Convert.ToDecimal(pctJuan);
                            else if (decimal.TryParse(pctJuan.ToString(), out decimal pj))
                                valor = pj;
                            
                            // Multiplicar por 100 para obtener el porcentaje entero
                            resumen.PorcentajeJuan = valor * 100;
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error al leer sueldos: {ex.Message}");
                }
            });
        }

        /// <summary>
        /// Guarda los sueldos en las columnas J y K del Excel (filas 7-8)
        /// </summary>
        public async Task<bool> GuardarSueldosAsync(DateTime fecha, decimal sueldoAndrea, decimal sueldoJuan)
        {
            return await Task.Run(() =>
            {
                try
                {
                    using (var package = new ExcelPackage(new FileInfo(_excelPath)))
                    {
                        var worksheet = BuscarHojaMes(package, fecha);
                        if (worksheet == null)
                            return false;

                        // Guardar sueldos en las celdas K7 y K8
                        worksheet.Cells[7, 11].Value = sueldoAndrea;
                        worksheet.Cells[8, 11].Value = sueldoJuan;

                        // Forzar recálculo de fórmulas
                        package.Workbook.Calculate();
                        // Marcar para recalcular al abrir en Excel
                        package.Workbook.CalcMode = ExcelCalcMode.Automatic;
                        package.Workbook.Properties.Modified = DateTime.Now;
                        
                        package.Save();
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error al guardar sueldos: {ex.Message}");
                    return false;
                }
            });
        }

        /// <summary>
        /// Crea todas las hojas faltantes desde la última existente hasta la fecha requerida
        /// </summary>
        private void CrearHojasFaltantes(ExcelPackage package, DateTime fechaDestino)
        {
            // Obtener la última hoja existente
            var ultimaHoja = package.Workbook.Worksheets
                .OrderByDescending(w => w.Name)
                .FirstOrDefault();
            
            DateTime fechaInicio;
            
            if (ultimaHoja != null)
            {
                // Intentar extraer la fecha de la última hoja
                fechaInicio = ExtraerFechaDeNombreHoja(ultimaHoja.Name);
                
                // Si no se pudo extraer, empezar desde el mes siguiente a la fecha actual
                if (fechaInicio == DateTime.MinValue)
                {
                    fechaInicio = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
                }
                else
                {
                    // Empezar desde el siguiente mes
                    fechaInicio = fechaInicio.AddMonths(1);
                }
            }
            else
            {
                // No hay hojas, crear desde la fecha destino
                fechaInicio = new DateTime(fechaDestino.Year, fechaDestino.Month, 1);
            }
            
            // Normalizar fechaDestino al primer día del mes
            DateTime fechaFin = new DateTime(fechaDestino.Year, fechaDestino.Month, 1);
            
            // Crear todas las hojas desde fechaInicio hasta fechaFin
            DateTime fechaActual = fechaInicio;
            while (fechaActual <= fechaFin)
            {
                string nombreHojaNueva = ObtenerNombreHoja(fechaActual);
                
                // Verificar si la hoja ya existe
                if (package.Workbook.Worksheets[nombreHojaNueva] == null)
                {
                    CrearHojaNueva(package, fechaActual);
                }
                
                fechaActual = fechaActual.AddMonths(1);
            }
        }
        
        /// <summary>
        /// Crea una nueva hoja copiando la plantilla de la última hoja existente
        /// </summary>
        private void CrearHojaNueva(ExcelPackage package, DateTime fecha)
        {
            string nombreHoja = ObtenerNombreHoja(fecha);
            
            // Buscar la última hoja existente para usar como plantilla
            var ultimaHoja = package.Workbook.Worksheets
                .OrderByDescending(w => w.Name)
                .FirstOrDefault();
            
            ExcelWorksheet worksheet;
            
            if (ultimaHoja != null)
            {
                // Copiar la última hoja completa (con fórmulas y formato)
                worksheet = package.Workbook.Worksheets.Add(nombreHoja, ultimaHoja);
                
                // Limpiar solo los datos de gastos (columnas A, B, C, D, E, H desde fila 2)
                if (worksheet.Dimension != null)
                {
                    int ultimaFilaCopiada = worksheet.Dimension.End.Row;
                    
                    // Borrar contenido de las columnas de gastos
                    for (int row = 2; row <= ultimaFilaCopiada; row++)
                    {
                        worksheet.Cells[row, 1].Value = null; // Fecha
                        worksheet.Cells[row, 2].Value = null; // Categoría
                        worksheet.Cells[row, 3].Value = null; // Monto
                        worksheet.Cells[row, 4].Value = null; // QuienPago
                        worksheet.Cells[row, 5].Value = null; // EsProporcional
                        worksheet.Cells[row, 8].Value = null; // Comentarios
                    }
                }
                
                // Actualizar fórmulas M2 y N2 con referencia al mes anterior
                try
                {
                    var mesAnterior = fecha.AddMonths(-1);
                    var nombreHojaAnterior = ObtenerNombreHoja(mesAnterior);
                    
                    // Verificar si la hoja anterior existe
                    var hojaAnteriorExiste = package.Workbook.Worksheets[nombreHojaAnterior] != null;
                    
                    if (hojaAnteriorExiste)
                    {
                        // Actualizar fórmula M2 (Debe Andrea)
                        worksheet.Cells["M2"].Formula = $"=SUM(F2:F50)-K2+'{nombreHojaAnterior}'!M2";
                        
                        // Actualizar fórmula N2 (Debe Juan)
                        worksheet.Cells["N2"].Formula = $"=SUM(G2:G50)-L2+'{nombreHojaAnterior}'!N2";
                    }
                    else
                    {
                        // Si es el primer mes, usar solo la suma sin referencia anterior
                        worksheet.Cells["M2"].Formula = "=SUM(F2:F50)-K2";
                        worksheet.Cells["N2"].Formula = "=SUM(G2:G50)-L2";
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"No se pudieron actualizar las fórmulas M2/N2: {ex.Message}");
                }
            }
            else
            {
                // Si no hay hojas existentes, crear una básica
                worksheet = package.Workbook.Worksheets.Add(nombreHoja);
                
                // Crear encabezados
                worksheet.Cells[1, 1].Value = "Fecha";
                worksheet.Cells[1, 2].Value = "Categoría";
                worksheet.Cells[1, 3].Value = "Monto";
                worksheet.Cells[1, 4].Value = "QuienPago";
                worksheet.Cells[1, 5].Value = "EsProporcional";
                worksheet.Cells[1, 8].Value = "Comentarios";
                
                // Aplicar estilo a encabezados
                using (var range = worksheet.Cells[1, 1, 1, 8])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(37, 99, 235));
                    range.Style.Font.Color.SetColor(System.Drawing.Color.White);
                }
                
                // Ajustar ancho de columnas
                worksheet.Column(1).Width = 12;
                worksheet.Column(2).Width = 20;
                worksheet.Column(3).Width = 12;
                worksheet.Column(4).Width = 15;
                worksheet.Column(5).Width = 15;
                worksheet.Column(8).Width = 35;
            }
        }
        
        /// <summary>
        /// Extrae la fecha de un nombre de hoja con formato "Mes-AA"
        /// </summary>
        private DateTime ExtraerFechaDeNombreHoja(string nombreHoja)
        {
            try
            {
                // Formato esperado: "Mayo-25" o "Diciembre-27"
                var partes = nombreHoja.Split('-');
                if (partes.Length != 2)
                    return DateTime.MinValue;
                
                string nombreMes = partes[0].Trim();
                string año = partes[1].Trim();
                
                // Convertir nombre del mes a número
                string[] meses = { "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
                                  "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre" };
                
                int mes = Array.IndexOf(meses, nombreMes) + 1;
                if (mes == 0)
                    return DateTime.MinValue;
                
                // Convertir año corto a año completo (asumiendo 20XX)
                int añoCompleto = 2000 + int.Parse(año);
                
                return new DateTime(añoCompleto, mes, 1);
            }
            catch
            {
                return DateTime.MinValue;
            }
        }

        public string RutaArchivo => _excelPath;

        public void Dispose()
        {
            // EPPlus no requiere limpieza de COM
        }
    }
}
