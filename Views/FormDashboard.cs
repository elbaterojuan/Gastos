using System;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using Gastos.Models;
using Gastos.Services;
using Gastos.Utils;

namespace Gastos.Views
{
    public partial class FormDashboard : Form
    {
        private ExcelService _excelService;
        private Panel panelHeader;
        private Panel panelStats;
        private Panel panelSueldos;
        private Panel panelCharts;
        private Label lblTitulo;
        private Label lblTotalGastos;
        private Label lblPromedioGastos;
        private Label lblCantidadGastos;
        private Label lblDeuda;
        private NumericUpDown txtSueldoAndrea;
        private NumericUpDown txtSueldoJuan;
        private Label lblSueldoTotal;
        private Label lblPorcentajeAndrea;
        private Label lblPorcentajeJuan;
        private Button btnGuardarSueldos;
        private Chart chartCategorias;
        private Chart chartPersonas;
        private ComboBox cboMes;
        private ComboBox cboAño;
        private Button btnActualizar;
        private Button btnVerDetalles;
        private Button btnMesAnterior;
        private Button btnMesSiguiente;

        public FormDashboard(ExcelService excelService)
        {
            _excelService = excelService;
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.WindowState = FormWindowState.Maximized;
            this.MinimumSize = new Size(1000, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "Dashboard de Gastos";
            this.BackColor = TemaColores.FondoClaro;
            this.Font = new Font("Segoe UI", 9.5F);
            this.AutoScroll = true;

            CrearPanelGraficos();  // Primero el Fill
            CrearPanelSueldos();  // Luego Top
            CrearPanelEstadisticas();  // Luego Top
            CrearHeader();  // Finalmente Top (queda arriba de todo)
            CargarDatosIniciales();
            
            // Establecer foco al inicio después de cargar
            this.Load += (s, e) => {
                this.AutoScrollPosition = new Point(0, 0);
                cboMes.Focus();
            };
        }

        private void CrearHeader()
        {
            panelHeader = new Panel
            {
                Dock = DockStyle.Top,
                Height = 80,
                BackColor = TemaColores.PrimarioAzul,
                Padding = new Padding(20)
            };

            lblTitulo = new Label
            {
                Text = "📊 Dashboard de Gastos",
                Font = new Font("Segoe UI", 18F, FontStyle.Bold),
                ForeColor = Color.White,
                AutoSize = true,
                Location = new Point(20, 25)
            };

            cboMes = new ComboBox
            {
                Width = 150,
                Location = new Point(350, 28),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Segoe UI", 11F),
                DropDownWidth = 150,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White
            };
            cboMes.Items.AddRange(new object[] {
                "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
                "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
            });
            cboMes.SelectedIndex = DateTime.Now.Month - 1;
            cboMes.SelectedIndexChanged += async (s, e) => await CargarDatos();

            cboAño = new ComboBox
            {
                Width = 100,
                Location = new Point(520, 28),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Segoe UI", 11F),
                DropDownWidth = 100,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White
            };
            // Agregar años desde 2020 hasta el año actual + 2
            int añoActual = DateTime.Now.Year;
            for (int año = 2020; año <= añoActual + 2; año++)
            {
                cboAño.Items.Add(año);
            }
            cboAño.SelectedItem = añoActual;
            cboAño.SelectedIndexChanged += async (s, e) => await CargarDatos();

            // Botón mes anterior
            btnMesAnterior = new Button
            {
                Text = "◀",
                Location = new Point(640, 25),
                Size = new Size(45, 35),
                BackColor = Color.White,
                ForeColor = TemaColores.PrimarioAzul,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 12F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnMesAnterior.FlatAppearance.BorderSize = 0;
            btnMesAnterior.Click += BtnMesAnterior_Click;

            // Botón mes siguiente
            btnMesSiguiente = new Button
            {
                Text = "▶",
                Location = new Point(690, 25),
                Size = new Size(45, 35),
                BackColor = Color.White,
                ForeColor = TemaColores.PrimarioAzul,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 12F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnMesSiguiente.FlatAppearance.BorderSize = 0;
            btnMesSiguiente.Click += BtnMesSiguiente_Click;

            btnActualizar = new Button
            {
                Text = "🔄 Actualizar",
                Location = new Point(760, 25),
                Size = new Size(120, 35),
                FlatStyle = FlatStyle.Flat,
                BackColor = TemaColores.SecundarioVerde,
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9.5F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnActualizar.FlatAppearance.BorderSize = 0;
            btnActualizar.Click += async (s, e) => await CargarDatos();

            btnVerDetalles = new Button
            {
                Text = "📋 Ver Detalles",
                Location = new Point(900, 25),
                Size = new Size(140, 35),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                ForeColor = TemaColores.PrimarioAzul,
                Font = new Font("Segoe UI", 9.5F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnVerDetalles.FlatAppearance.BorderSize = 0;
            btnVerDetalles.Click += BtnVerDetalles_Click;

            panelHeader.Controls.AddRange(new Control[] { lblTitulo, cboMes, cboAño, btnMesAnterior, btnMesSiguiente, btnActualizar, btnVerDetalles });
            this.Controls.Add(panelHeader);
        }

        private void CrearPanelEstadisticas()
        {
            panelStats = new Panel
            {
                Dock = DockStyle.Top,
                Height = 120,
                BackColor = TemaColores.FondoClaro,
                Padding = new Padding(20, 20, 20, 10)
            };

            // Panel Total
            var panelTotal = CrearPanelStat("💰 Total Gastos", "$0.00", TemaColores.PrimarioAzul, 0);
            lblTotalGastos = (Label)panelTotal.Controls[1];

            // Panel Promedio
            var panelPromedio = CrearPanelStat("📊 Promedio", "$0.00", TemaColores.SecundarioVerde, 230);
            lblPromedioGastos = (Label)panelPromedio.Controls[1];

            // Panel Cantidad
            var panelCantidad = CrearPanelStat("🔢 Cantidad", "0", TemaColores.AccentoNaranja, 460);
            lblCantidadGastos = (Label)panelCantidad.Controls[1];

            // Panel Deuda
            var panelDeuda = CrearPanelStat("💳 Deuda", "$0.00", Color.FromArgb(239, 68, 68), 690);
            lblDeuda = (Label)panelDeuda.Controls[1];

            panelStats.Controls.AddRange(new Control[] { panelTotal, panelPromedio, panelCantidad, panelDeuda });
            this.Controls.Add(panelStats);
        }

        private void CrearPanelSueldos()
        {
            panelSueldos = new Panel
            {
                Dock = DockStyle.Top,
                Height = 100,
                BackColor = Color.White,
                Padding = new Padding(20)
            };

            var lblTitulo = new Label
            {
                Text = "💼 Sueldos y Porcentajes",
                Font = new Font("Segoe UI", 12F, FontStyle.Bold),
                ForeColor = TemaColores.TextoOscuro,
                AutoSize = true,
                Location = new Point(20, 15)
            };

            // Sueldo Andrea
            var lblAndrea = new Label
            {
                Text = "Sueldo Andrea:",
                Location = new Point(20, 50),
                AutoSize = true,
                Font = new Font("Segoe UI", 9.5F)
            };

            txtSueldoAndrea = new NumericUpDown
            {
                Location = new Point(130, 47),
                Width = 120,
                Maximum = 999999999,
                DecimalPlaces = 2,
                ThousandsSeparator = true,
                Font = new Font("Segoe UI", 10F)
            };
            txtSueldoAndrea.ValueChanged += (s, e) => ActualizarCalculosSueldos();

            // Sueldo Juan
            var lblJuan = new Label
            {
                Text = "Sueldo Juan:",
                Location = new Point(270, 50),
                AutoSize = true,
                Font = new Font("Segoe UI", 9.5F)
            };

            txtSueldoJuan = new NumericUpDown
            {
                Location = new Point(365, 47),
                Width = 120,
                Maximum = 999999999,
                DecimalPlaces = 2,
                ThousandsSeparator = true,
                Font = new Font("Segoe UI", 10F)
            };
            txtSueldoJuan.ValueChanged += (s, e) => ActualizarCalculosSueldos();

            // Total
            var lblTotalTexto = new Label
            {
                Text = "Total:",
                Location = new Point(510, 50),
                AutoSize = true,
                Font = new Font("Segoe UI", 9.5F, FontStyle.Bold)
            };

            lblSueldoTotal = new Label
            {
                Text = "$0.00",
                Location = new Point(565, 50),
                AutoSize = true,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                ForeColor = TemaColores.PrimarioAzul
            };

            // Porcentaje Andrea
            var lblPctAndreaTexto = new Label
            {
                Text = "% Andrea:",
                Location = new Point(650, 50),
                AutoSize = true,
                Font = new Font("Segoe UI", 9.5F)
            };

            lblPorcentajeAndrea = new Label
            {
                Text = "0%",
                Location = new Point(725, 50),
                AutoSize = true,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                ForeColor = TemaColores.SecundarioVerde
            };

            // Porcentaje Juan
            var lblPctJuanTexto = new Label
            {
                Text = "% Juan:",
                Location = new Point(780, 50),
                AutoSize = true,
                Font = new Font("Segoe UI", 9.5F)
            };

            lblPorcentajeJuan = new Label
            {
                Text = "0%",
                Location = new Point(840, 50),
                AutoSize = true,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                ForeColor = TemaColores.AccentoNaranja
            };

            // Botón Guardar
            btnGuardarSueldos = new Button
            {
                Text = "💾 Guardar",
                Location = new Point(920, 45),
                Size = new Size(100, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = TemaColores.SecundarioVerde,
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9.5F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnGuardarSueldos.FlatAppearance.BorderSize = 0;
            btnGuardarSueldos.Click += BtnGuardarSueldos_Click;

            panelSueldos.Controls.AddRange(new Control[] {
                lblTitulo, lblAndrea, txtSueldoAndrea, lblJuan, txtSueldoJuan,
                lblTotalTexto, lblSueldoTotal, lblPctAndreaTexto, lblPorcentajeAndrea,
                lblPctJuanTexto, lblPorcentajeJuan, btnGuardarSueldos
            });

            this.Controls.Add(panelSueldos);
        }

        private Panel CrearPanelStat(string titulo, string valor, Color color, int x)
        {
            var panel = new Panel
            {
                Size = new Size(220, 90),
                Location = new Point(x, 0),
                BackColor = Color.White
            };
            panel.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, panel.ClientRectangle,
                TemaColores.BordeGris, ButtonBorderStyle.Solid);

            var lblTitulo = new Label
            {
                Text = titulo,
                Font = new Font("Segoe UI", 10F),
                ForeColor = TemaColores.TextoGris,
                AutoSize = true,
                Location = new Point(15, 15)
            };

            var lblValor = new Label
            {
                Text = valor,
                Font = new Font("Segoe UI", 20F, FontStyle.Bold),
                ForeColor = color,
                AutoSize = true,
                Location = new Point(15, 40)
            };

            panel.Controls.AddRange(new Control[] { lblTitulo, lblValor });
            return panel;
        }

        private void CrearPanelGraficos()
        {
            panelCharts = new Panel
            {
                Dock = DockStyle.Top,
                Height = 450,
                BackColor = TemaColores.FondoClaro,
                Padding = new Padding(20)
            };

            // Gráfico por Categorías
            chartCategorias = CrearGrafico("Gastos por Categoría", 0);
            
            // Gráfico por Personas
            chartPersonas = CrearGrafico("Gastos por Persona", 600);

            panelCharts.Controls.AddRange(new Control[] { chartCategorias, chartPersonas });
            this.Controls.Add(panelCharts);
        }

        private Chart CrearGrafico(string titulo, int x)
        {
            var chart = new Chart
            {
                Size = new Size(550, 450),
                Location = new Point(x, 0),
                BackColor = Color.White
            };

            var chartArea = new ChartArea
            {
                BackColor = Color.White,
                BorderColor = TemaColores.BordeGris,
                BorderWidth = 1
            };
            chart.ChartAreas.Add(chartArea);

            var legend = new Legend
            {
                Docking = Docking.Bottom,
                Font = new Font("Segoe UI", 9F),
                IsTextAutoFit = true
            };
            chart.Legends.Add(legend);

            var title = new Title
            {
                Text = titulo,
                Font = new Font("Segoe UI", 12F, FontStyle.Bold),
                ForeColor = TemaColores.TextoOscuro
            };
            chart.Titles.Add(title);

            var series = new Series
            {
                ChartType = SeriesChartType.Pie,
                Font = new Font("Segoe UI", 9F),
                IsValueShownAsLabel = true,
                LabelFormat = "${0:N0}"
            };
            chart.Series.Add(series);

            return chart;
        }

        private async void CargarDatosIniciales()
        {
            await CargarDatos();
        }

        private async Task CargarDatos()
        {
            try
            {
                btnActualizar.Enabled = false;
                btnActualizar.Text = "⏳ Cargando...";

                int añoSeleccionado = cboAño.SelectedItem != null ? (int)cboAño.SelectedItem : DateTime.Now.Year;
                var fecha = new DateTime(añoSeleccionado, cboMes.SelectedIndex + 1, 1);
                var resumen = await _excelService.ObtenerResumenMensualAsync(fecha);

                ActualizarEstadisticas(resumen);
                ActualizarGraficos(resumen);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al cargar datos: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnActualizar.Enabled = true;
                btnActualizar.Text = "🔄 Actualizar";
            }
        }

        private void ActualizarEstadisticas(ResumenMensual resumen)
        {
            lblTotalGastos.Text = $"${resumen.TotalGastado:N2}";
            lblPromedioGastos.Text = $"${resumen.PromedioGasto:N2}";
            lblCantidadGastos.Text = resumen.CantidadGastos.ToString();
            
            // Actualizar panel de deuda
            if (!string.IsNullOrEmpty(resumen.DeudorNombre) && resumen.DeudorMonto > 0)
            {
                // Actualizar título del panel de deuda
                var panelDeuda = lblDeuda.Parent;
                var lblTituloDeuda = panelDeuda.Controls[0] as Label;
                if (lblTituloDeuda != null)
                {
                    lblTituloDeuda.Text = $"💳 Debe {resumen.DeudorNombre}";
                }
                lblDeuda.Text = $"${resumen.DeudorMonto:N2}";
                
                // Cambiar color del monto según quién debe
                if (resumen.DeudorNombre == "Andrea")
                {
                    lblDeuda.ForeColor = Color.DeepPink;
                }
                else if (resumen.DeudorNombre == "Juan")
                {
                    lblDeuda.ForeColor = Color.FromArgb(59, 130, 246); // Azul
                }
            }
            else
            {
                lblDeuda.Text = "$0.00";
                lblDeuda.ForeColor = Color.White;
            }

            // Actualizar sueldos
            txtSueldoAndrea.Value = resumen.SueldoAndrea;
            txtSueldoJuan.Value = resumen.SueldoJuan;
            lblSueldoTotal.Text = $"${resumen.SueldoTotal:N2}";
            lblPorcentajeAndrea.Text = $"{resumen.PorcentajeAndrea:N0}%";
            lblPorcentajeJuan.Text = $"{resumen.PorcentajeJuan:N0}%";
        }

        private void ActualizarCalculosSueldos()
        {
            decimal total = txtSueldoAndrea.Value + txtSueldoJuan.Value;
            lblSueldoTotal.Text = $"${total:N2}";

            if (total > 0)
            {
                decimal pctAndrea = (txtSueldoAndrea.Value / total) * 100;
                decimal pctJuan = (txtSueldoJuan.Value / total) * 100;
                lblPorcentajeAndrea.Text = $"{pctAndrea:N0}%";
                lblPorcentajeJuan.Text = $"{pctJuan:N0}%";
            }
            else
            {
                lblPorcentajeAndrea.Text = "0%";
                lblPorcentajeJuan.Text = "0%";
            }
        }

        private async void BtnGuardarSueldos_Click(object sender, EventArgs e)
        {
            try
            {
                btnGuardarSueldos.Enabled = false;
                btnGuardarSueldos.Text = "⏳ Guardando...";

                int añoSeleccionado = cboAño.SelectedItem != null ? (int)cboAño.SelectedItem : DateTime.Now.Year;
                var fecha = new DateTime(añoSeleccionado, cboMes.SelectedIndex + 1, 1);

                bool resultado = await _excelService.GuardarSueldosAsync(fecha, txtSueldoAndrea.Value, txtSueldoJuan.Value);

                if (resultado)
                {
                    MessageBox.Show("Sueldos guardados correctamente", "Éxito",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Error al guardar sueldos", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al guardar sueldos: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnGuardarSueldos.Enabled = true;
                btnGuardarSueldos.Text = "💾 Guardar";
            }
        }

        private void ActualizarGraficos(ResumenMensual resumen)
        {
            // Gráfico por Categorías
            chartCategorias.Series[0].Points.Clear();
            if (resumen.GastosPorCategoria.Any())
            {
                var colores = new[] { "#2563EB", "#10B981", "#F59E0B", "#EF4444", "#8B5CF6", "#EC4899" };
                int i = 0;
                foreach (var item in resumen.GastosPorCategoria.OrderByDescending(x => x.Value))
                {
                    var point = chartCategorias.Series[0].Points.AddXY(item.Key, item.Value);
                    chartCategorias.Series[0].Points[point].Color = ColorTranslator.FromHtml(colores[i % colores.Length]);
                    i++;
                }
            }

            // Gráfico por Personas
            chartPersonas.Series[0].Points.Clear();
            if (resumen.GastosPorPersona.Any())
            {
                foreach (var item in resumen.GastosPorPersona)
                {
                    chartPersonas.Series[0].Points.AddXY(item.Key, item.Value);
                }
            }
        }

        private async void BtnMesAnterior_Click(object sender, EventArgs e)
        {
            if (cboMes.SelectedIndex > 0)
            {
                cboMes.SelectedIndex--;
            }
            else
            {
                cboMes.SelectedIndex = 11;
                if (cboAño.SelectedIndex > 0)
                {
                    cboAño.SelectedIndex--;
                }
            }
        }

        private async void BtnMesSiguiente_Click(object sender, EventArgs e)
        {
            if (cboMes.SelectedIndex < 11)
            {
                cboMes.SelectedIndex++;
            }
            else
            {
                cboMes.SelectedIndex = 0;
                if (cboAño.SelectedIndex < cboAño.Items.Count - 1)
                {
                    cboAño.SelectedIndex++;
                }
            }
        }

        private void BtnVerDetalles_Click(object sender, EventArgs e)
        {
            int añoSeleccionado = cboAño.SelectedItem != null ? (int)cboAño.SelectedItem : DateTime.Now.Year;
            var fecha = new DateTime(añoSeleccionado, cboMes.SelectedIndex + 1, 1);
            var formDetalles = new FormDetalles(_excelService, fecha);
            formDetalles.ShowDialog();
        }
    }
}
