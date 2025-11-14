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
        private Panel panelCharts;
        private Label lblTitulo;
        private Label lblTotalGastos;
        private Label lblPromedioGastos;
        private Label lblCantidadGastos;
        private Chart chartCategorias;
        private Chart chartPersonas;
        private ComboBox cboMes;
        private Button btnActualizar;
        private Button btnVerDetalles;

        public FormDashboard(ExcelService excelService)
        {
            _excelService = excelService;
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Size = new Size(1200, 700);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "Dashboard de Gastos";
            this.BackColor = TemaColores.FondoClaro;
            this.Font = new Font("Segoe UI", 9.5F);

            CrearHeader();
            CrearPanelEstadisticas();
            CrearPanelGraficos();
            CargarDatosIniciales();
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
                Width = 200,
                Location = new Point(300, 28),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Segoe UI", 10F)
            };
            cboMes.Items.AddRange(new object[] {
                "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
                "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
            });
            cboMes.SelectedIndex = DateTime.Now.Month - 1;
            cboMes.SelectedIndexChanged += async (s, e) => await CargarDatos();

            btnActualizar = new Button
            {
                Text = "🔄 Actualizar",
                Location = new Point(520, 25),
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
                Location = new Point(660, 25),
                Size = new Size(140, 35),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                ForeColor = TemaColores.PrimarioAzul,
                Font = new Font("Segoe UI", 9.5F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnVerDetalles.FlatAppearance.BorderSize = 0;
            btnVerDetalles.Click += BtnVerDetalles_Click;

            panelHeader.Controls.AddRange(new Control[] { lblTitulo, cboMes, btnActualizar, btnVerDetalles });
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
            var panelPromedio = CrearPanelStat("📊 Promedio", "$0.00", TemaColores.SecundarioVerde, 300);
            lblPromedioGastos = (Label)panelPromedio.Controls[1];

            // Panel Cantidad
            var panelCantidad = CrearPanelStat("🔢 Cantidad", "0", TemaColores.AccentoNaranja, 600);
            lblCantidadGastos = (Label)panelCantidad.Controls[1];

            panelStats.Controls.AddRange(new Control[] { panelTotal, panelPromedio, panelCantidad });
            this.Controls.Add(panelStats);
        }

        private Panel CrearPanelStat(string titulo, string valor, Color color, int x)
        {
            var panel = new Panel
            {
                Size = new Size(280, 90),
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
                Dock = DockStyle.Fill,
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

                var fecha = new DateTime(DateTime.Now.Year, cboMes.SelectedIndex + 1, 1);
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
            lblTotalGastos.Text = $"${resumen.TotalGastos:N2}";
            lblPromedioGastos.Text = $"${resumen.PromedioGastos:N2}";
            lblCantidadGastos.Text = resumen.CantidadGastos.ToString();
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

        private void BtnVerDetalles_Click(object sender, EventArgs e)
        {
            var fecha = new DateTime(DateTime.Now.Year, cboMes.SelectedIndex + 1, 1);
            var formDetalles = new FormDetalles(_excelService, fecha);
            formDetalles.ShowDialog();
        }
    }
}
