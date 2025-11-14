using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Gastos.Models;
using Gastos.Services;
using Gastos.Utils;

namespace Gastos.Views
{
    public partial class FormDetalles : Form
    {
        private ExcelService _excelService;
        private DateTime _fecha;
        private List<Gasto> _gastos;
        
        private Panel panelHeader;
        private Panel panelFiltros;
        private DataGridView dgvGastos;
        private TextBox txtBuscar;
        private ComboBox cboCategoria;
        private Button btnExportar;
        private Button btnMesAnterior;
        private Button btnMesSiguiente;
        private Label lblTitulo;
        private Label lblTotal;

        public FormDetalles(ExcelService excelService, DateTime fecha)
        {
            _excelService = excelService;
            _fecha = fecha;
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.WindowState = FormWindowState.Maximized;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = $"Detalle de Gastos - {_fecha:MMMM yyyy}";
            this.BackColor = TemaColores.FondoClaro;
            this.Font = new Font("Segoe UI", 9.5F);

            CrearDataGridView();  // Primero el Fill
            CrearPanelFiltros();  // Luego Top
            CrearHeader();  // Finalmente Top (queda arriba de todo)
            CargarDatos();
        }

        private void CrearHeader()
        {
            panelHeader = new Panel
            {
                Dock = DockStyle.Top,
                Height = 70,
                BackColor = TemaColores.PrimarioAzul,
                Padding = new Padding(20)
            };

            lblTitulo = new Label
            {
                Text = $"📋 Detalle de Gastos - {_fecha:MMMM yyyy}",
                Font = new Font("Segoe UI", 16F, FontStyle.Bold),
                ForeColor = Color.White,
                AutoSize = true,
                Location = new Point(20, 20)
            };

            // Botón mes anterior
            btnMesAnterior = new Button
            {
                Text = "◀",
                Location = new Point(500, 18),
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
                Location = new Point(550, 18),
                Size = new Size(45, 35),
                BackColor = Color.White,
                ForeColor = TemaColores.PrimarioAzul,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 12F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnMesSiguiente.FlatAppearance.BorderSize = 0;
            btnMesSiguiente.Click += BtnMesSiguiente_Click;

            lblTotal = new Label
            {
                Text = "Total: $0.00",
                Font = new Font("Segoe UI", 14F, FontStyle.Bold),
                ForeColor = Color.White,
                AutoSize = true,
                Location = new Point(650, 22)
            };

            panelHeader.Controls.AddRange(new Control[] { lblTitulo, btnMesAnterior, btnMesSiguiente, lblTotal });
            this.Controls.Add(panelHeader);
        }

        private void CrearPanelFiltros()
        {
            panelFiltros = new Panel
            {
                Dock = DockStyle.Top,
                Height = 70,
                BackColor = Color.White,
                Padding = new Padding(20, 15, 20, 15)
            };

            var lblBuscar = new Label
            {
                Text = "🔍 Buscar:",
                AutoSize = true,
                Location = new Point(20, 22),
                Font = new Font("Segoe UI", 9.5F)
            };

            txtBuscar = new TextBox
            {
                Width = 250,
                Location = new Point(95, 18)
            };
            txtBuscar.AplicarEstiloTextBox();
            txtBuscar.TextChanged += (s, e) => FiltrarGastos();

            var lblCategoria = new Label
            {
                Text = "Categoría:",
                AutoSize = true,
                Location = new Point(370, 22),
                Font = new Font("Segoe UI", 9.5F)
            };

            cboCategoria = new ComboBox
            {
                Width = 200,
                Location = new Point(450, 18),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cboCategoria.AplicarEstiloComboBox();
            cboCategoria.Items.Add("Todas");
            cboCategoria.SelectedIndex = 0;
            cboCategoria.SelectedIndexChanged += (s, e) => FiltrarGastos();

            btnExportar = new Button
            {
                Text = "📄 Exportar PDF",
                Location = new Point(680, 15),
                Size = new Size(140, 38)
            };
            btnExportar.AplicarEstiloBotonSecundario();
            btnExportar.Click += BtnExportar_Click;

            panelFiltros.Controls.AddRange(new Control[] { lblBuscar, txtBuscar, lblCategoria, cboCategoria, btnExportar });
            this.Controls.Add(panelFiltros);
        }

        private void CrearDataGridView()
        {
            dgvGastos = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoGenerateColumns = false
            };
            dgvGastos.AplicarEstiloDataGridView();

            // Configurar columnas
            dgvGastos.Columns.AddRange(new DataGridViewColumn[]
            {
                new DataGridViewTextBoxColumn
                {
                    DataPropertyName = "Fecha",
                    HeaderText = "📅 Fecha",
                    Width = 100,
                    DefaultCellStyle = new DataGridViewCellStyle { Format = "dd/MM/yyyy" }
                },
                new DataGridViewTextBoxColumn
                {
                    DataPropertyName = "Categoria",
                    HeaderText = "🏷️ Categoría",
                    Width = 150
                },
                new DataGridViewTextBoxColumn
                {
                    DataPropertyName = "Monto",
                    HeaderText = "💰 Monto",
                    Width = 120,
                    DefaultCellStyle = new DataGridViewCellStyle 
                    { 
                        Format = "$#,##0.00",
                        Alignment = DataGridViewContentAlignment.MiddleRight
                    }
                },
                new DataGridViewTextBoxColumn
                {
                    DataPropertyName = "QuienPago",
                    HeaderText = "👤 Quién Pagó",
                    Width = 130
                },
                new DataGridViewCheckBoxColumn
                {
                    DataPropertyName = "EsProporcional",
                    HeaderText = "� Proporcional",
                    Width = 110
                },
                new DataGridViewTextBoxColumn
                {
                    DataPropertyName = "Comentarios",
                    HeaderText = "📝 Comentarios",
                    AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                }
            });

            this.Controls.Add(dgvGastos);
        }

        private async void CargarDatos()
        {
            try
            {
                _gastos = await _excelService.ObtenerGastosMesAsync(_fecha);
                
                // Cargar categorías únicas
                var categorias = _gastos.Select(g => g.Categoria).Distinct().OrderBy(c => c);
                foreach (var cat in categorias)
                {
                    if (!cboCategoria.Items.Contains(cat))
                        cboCategoria.Items.Add(cat);
                }

                MostrarGastos(_gastos);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al cargar gastos: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void FiltrarGastos()
        {
            if (_gastos == null) return;

            var gastosFiltrados = _gastos.AsEnumerable();

            // Filtro por búsqueda
            if (!string.IsNullOrWhiteSpace(txtBuscar.Text))
            {
                var busqueda = txtBuscar.Text.ToLower();
                gastosFiltrados = gastosFiltrados.Where(g =>
                    g.Categoria.ToLower().Contains(busqueda) ||
                    g.Comentarios.ToLower().Contains(busqueda) ||
                    g.QuienPago.ToLower().Contains(busqueda));
            }

            // Filtro por categoría
            if (cboCategoria.SelectedIndex > 0)
            {
                gastosFiltrados = gastosFiltrados.Where(g => g.Categoria == cboCategoria.Text);
            }

            MostrarGastos(gastosFiltrados.ToList());
        }

        private void MostrarGastos(List<Gasto> gastos)
        {
            dgvGastos.DataSource = null;
            dgvGastos.DataSource = gastos;

            var total = gastos.Sum(g => g.Monto);
            lblTotal.Text = $"Total: ${total:N2}";
        }

        private async void BtnMesAnterior_Click(object sender, EventArgs e)
        {
            _fecha = _fecha.AddMonths(-1);
            lblTitulo.Text = $"📋 Detalle de Gastos - {_fecha:MMMM yyyy}";
            this.Text = $"Detalle de Gastos - {_fecha:MMMM yyyy}";
            CargarDatos();
        }

        private async void BtnMesSiguiente_Click(object sender, EventArgs e)
        {
            _fecha = _fecha.AddMonths(1);
            lblTitulo.Text = $"📋 Detalle de Gastos - {_fecha:MMMM yyyy}";
            this.Text = $"Detalle de Gastos - {_fecha:MMMM yyyy}";
            CargarDatos();
        }

        private void BtnExportar_Click(object sender, EventArgs e)
        {
            try
            {
                var saveDialog = new SaveFileDialog
                {
                    Filter = "PDF files (*.pdf)|*.pdf",
                    FileName = $"Gastos_{_fecha:yyyy-MM}.pdf",
                    Title = "Exportar a PDF"
                };

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    // Aquí iría la lógica de exportación a PDF
                    // Por ahora solo mostramos un mensaje
                    MessageBox.Show("Funcionalidad de exportación a PDF en desarrollo", "Información",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al exportar: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
