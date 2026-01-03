using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Net;
using Gastos.Models;
using Gastos.Services;
using Gastos.Utils;

namespace Gastos.Views
{
    public partial class FormDetalles : Form
    {
        private ExcelService _excelService;
        private DateTime _fecha;
        private string _nombreHoja; // Nombre de la hoja del mes
        private List<Gasto> _gastos;
        private List<Gasto> _gastosOriginales; // Copia de respaldo
        private Gasto _gastoBeingEdited;
        
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
            
            // Calcular el nombre de la hoja según el mes
            string[] meses = { "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
                              "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre" };
            string nombreMes = meses[fecha.Month - 1];
            string año = fecha.ToString("yy");
            _nombreHoja = $"{nombreMes}-{año}";
            
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.WindowState = FormWindowState.Maximized;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = $"Detalle de Gastos - {_fecha:MMMM yyyy}";
            this.BackColor = TemaColores.FondoClaro;
            this.Font = new Font("Segoe UI", 9.5F);
            this.FormClosing += FormDetalles_FormClosing;

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
                AutoGenerateColumns = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                MultiSelect = false,
                EditMode = DataGridViewEditMode.EditOnEnter,
                ReadOnly = false
            };
            dgvGastos.AplicarEstiloDataGridView();
            
            // Asegurar que sea editable después de aplicar estilos
            dgvGastos.ReadOnly = false;
            
            // Eventos para edición y eliminación
            dgvGastos.CellBeginEdit += DgvGastos_CellBeginEdit;
            dgvGastos.CellEndEdit += DgvGastos_CellEndEdit;
            dgvGastos.CellContentClick += DgvGastos_CellContentClick;
            dgvGastos.CellPainting += DgvGastos_CellPainting;

            // Configurar columnas
            dgvGastos.Columns.AddRange(new DataGridViewColumn[]
            {
                new DataGridViewTextBoxColumn
                {
                    DataPropertyName = "Fecha",
                    HeaderText = "📅 Fecha",
                    Width = 100,
                    ReadOnly = false,
                    DefaultCellStyle = new DataGridViewCellStyle { Format = "dd/MM/yyyy" }
                },
                new DataGridViewComboBoxColumn
                {
                    DataPropertyName = "Categoria",
                    HeaderText = "🏷️ Categoría",
                    Width = 150,
                    ReadOnly = false,
                    DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing
                },
                new DataGridViewTextBoxColumn
                {
                    DataPropertyName = "Monto",
                    HeaderText = "💰 Monto",
                    Width = 120,
                    ReadOnly = false,
                    DefaultCellStyle = new DataGridViewCellStyle 
                    { 
                        Format = "$#,##0.00",
                        Alignment = DataGridViewContentAlignment.MiddleRight
                    }
                },
                new DataGridViewComboBoxColumn
                {
                    DataPropertyName = "QuienPago",
                    HeaderText = "👤 Quién Pagó",
                    Width = 130,
                    ReadOnly = false,
                    DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing
                },
                new DataGridViewCheckBoxColumn
                {
                    DataPropertyName = "EsProporcional",
                    HeaderText = "� Proporcional",
                    Width = 110,
                    ReadOnly = false
                },
                new DataGridViewTextBoxColumn
                {
                    DataPropertyName = "Comentarios",
                    HeaderText = "📝 Comentarios",
                    AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                    ReadOnly = false
                },
                new DataGridViewButtonColumn
                {
                    HeaderText = "",
                    Name = "btnEliminar",
                    Width = 40,
                    Text = "✖",
                    UseColumnTextForButtonValue = false
                }
            });

            this.Controls.Add(dgvGastos);
        }

        private async void CargarDatos()
        {
            try
            {
                _gastos = await _excelService.ObtenerGastosMesAsync(_fecha);

                // Leer categorías desde app settings (App.config)
                var categoriasDesdeSettings = new List<string>();
                try
                {
                    var raw = Properties.Settings.Default.Categorias ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(raw))
                    {
                        var decoded = WebUtility.HtmlDecode(raw).Trim();
                        if (!string.IsNullOrWhiteSpace(decoded) && decoded.StartsWith("<"))
                        {
                            var doc = XDocument.Parse(decoded);
                            categoriasDesdeSettings = doc.Descendants("string")
                                .Select(x => x.Value)
                                .Where(s => !string.IsNullOrWhiteSpace(s))
                                .Distinct()
                                .ToList();
                        }
                    }
                }
                catch { /* ignorar errores de parseo de settings */ }

                // Cargar categorías únicas desde gastos y unir con las definidas en settings
                var categoriasDesdeGastos = _gastos.Select(g => g.Categoria)
                    .Where(c => !string.IsNullOrWhiteSpace(c))
                    .Distinct()
                    .ToList();

                var categorias = categoriasDesdeSettings
                    .Union(categoriasDesdeGastos)
                    .Distinct()
                    .OrderBy(c => c)
                    .ToList();

                // Actualizar cboCategoria
                cboCategoria.Items.Clear();
                cboCategoria.Items.Add("Todas");
                foreach (var cat in categorias)
                {
                    cboCategoria.Items.Add(cat);
                }
                cboCategoria.SelectedIndex = 0;

                // Configurar ComboBox de Categoría en el DataGridView
                var colCategoria = (DataGridViewComboBoxColumn)dgvGastos.Columns[1];
                colCategoria.Items.Clear();
                foreach (var cat in categorias)
                    colCategoria.Items.Add(cat);

                // Configurar ComboBox de QuienPago en el DataGridView
                var colQuienPago = (DataGridViewComboBoxColumn)dgvGastos.Columns[3];
                colQuienPago.Items.Clear();
                colQuienPago.Items.Add("Andrea");
                colQuienPago.Items.Add("Juan");

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

            // Crear una copia profunda de los gastos originales para comparación
            _gastosOriginales = gastos.Select(g => new Gasto
            {
                Fecha = g.Fecha,
                Categoria = g.Categoria,
                Monto = g.Monto,
                QuienPago = g.QuienPago,
                EsProporcional = g.EsProporcional,
                Comentarios = g.Comentarios
            }).ToList();

            var total = gastos.Sum(g => g.Monto);
            lblTotal.Text = $"Total: ${total:N2}";
        }

        private async void BtnMesAnterior_Click(object sender, EventArgs e)
        {
            _fecha = _fecha.AddMonths(-1);
            
            // Actualizar nombre de la hoja
            string[] meses = { "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
                              "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre" };
            string nombreMes = meses[_fecha.Month - 1];
            string año = _fecha.ToString("yy");
            _nombreHoja = $"{nombreMes}-{año}";
            
            lblTitulo.Text = $"📋 Detalle de Gastos - {_fecha:MMMM yyyy}";
            this.Text = $"Detalle de Gastos - {_fecha:MMMM yyyy}";
            CargarDatos();
        }

        private async void BtnMesSiguiente_Click(object sender, EventArgs e)
        {
            _fecha = _fecha.AddMonths(1);
            
            // Actualizar nombre de la hoja
            string[] meses = { "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
                              "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre" };
            string nombreMes = meses[_fecha.Month - 1];
            string año = _fecha.ToString("yy");
            _nombreHoja = $"{nombreMes}-{año}";
            
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

        private async void FormDetalles_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Forzar que termine la edición antes de cerrar
            if (dgvGastos.IsCurrentCellInEditMode && _gastoBeingEdited != null)
            {
                dgvGastos.EndEdit();
                
                // Dar tiempo para que se complete el guardado asíncrono
                if (dgvGastos.CurrentCell != null && dgvGastos.CurrentRow != null)
                {
                    var rowIndex = dgvGastos.CurrentCell.RowIndex;
                    if (rowIndex >= 0 && rowIndex < _gastos.Count)
                    {
                        try
                        {
                            var gastoActualizado = (Gasto)dgvGastos.Rows[rowIndex].DataBoundItem;
                            await _excelService.ActualizarGastoAsync(_gastoBeingEdited, gastoActualizado, _nombreHoja);
                            _gastoBeingEdited = null;
                        }
                        catch { }
                    }
                }
            }
        }

        private void DgvGastos_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            // Capturar el gasto ANTES de que se modifique desde la lista de respaldo
            if (e.RowIndex >= 0 && _gastosOriginales != null && e.RowIndex < _gastosOriginales.Count)
            {
                var original = _gastosOriginales[e.RowIndex];
                if (original != null)
                {
                    _gastoBeingEdited = new Gasto
                    {
                        Fecha = original.Fecha,
                        Categoria = original.Categoria,
                        Monto = original.Monto,
                        QuienPago = original.QuienPago,
                        EsProporcional = original.EsProporcional,
                        Comentarios = original.Comentarios
                    };
                }
            }
        }

        private async void DgvGastos_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.RowIndex < 0 || e.ColumnIndex == 6 || _gastoBeingEdited == null)
                    return;

                var gastoActualizado = (Gasto)dgvGastos.Rows[e.RowIndex].DataBoundItem;

                // Verificar si realmente hubo cambios
                bool huboChangios = _gastoBeingEdited.Fecha != gastoActualizado.Fecha ||
                                   _gastoBeingEdited.Categoria != gastoActualizado.Categoria ||
                                   _gastoBeingEdited.Monto != gastoActualizado.Monto ||
                                   _gastoBeingEdited.QuienPago != gastoActualizado.QuienPago ||
                                   _gastoBeingEdited.EsProporcional != gastoActualizado.EsProporcional ||
                                   _gastoBeingEdited.Comentarios != gastoActualizado.Comentarios;

                if (!huboChangios)
                {
                    _gastoBeingEdited = null;
                    return; // No hay cambios, no hacer nada
                }

                bool resultado = await _excelService.ActualizarGastoAsync(_gastoBeingEdited, gastoActualizado, _nombreHoja);

                if (resultado)
                {
                    // Actualizar la lista de respaldo
                    if (_gastosOriginales != null && e.RowIndex < _gastosOriginales.Count)
                    {
                        _gastosOriginales[e.RowIndex] = new Gasto
                        {
                            Fecha = gastoActualizado.Fecha,
                            Categoria = gastoActualizado.Categoria,
                            Monto = gastoActualizado.Monto,
                            QuienPago = gastoActualizado.QuienPago,
                            EsProporcional = gastoActualizado.EsProporcional,
                            Comentarios = gastoActualizado.Comentarios
                        };
                    }
                    
                    var total = _gastos.Sum(g => g.Monto);
                    lblTotal.Text = $"Total: ${total:N2}";
                    _gastoBeingEdited = null;
                }
                else
                {
                    MessageBox.Show("No se pudo actualizar el gasto en Excel", "Error");
                    _gastoBeingEdited = null;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error");
                _gastoBeingEdited = null;
            }
        }

        private async void DgvGastos_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            // Verificar si es la columna del botón eliminar (ahora es la última columna, índice 6)
            if (e.ColumnIndex == 6 && e.RowIndex >= 0)
            {
                try
                {
                    var gasto = (Gasto)dgvGastos.Rows[e.RowIndex].DataBoundItem;

                    var confirmacion = MessageBox.Show(
                        $"¿Está seguro de eliminar el gasto?\n\n" +
                        $"Fecha: {gasto.Fecha:dd/MM/yyyy}\n" +
                        $"Categoría: {gasto.Categoria}\n" +
                        $"Monto: ${gasto.Monto:N2}",
                        "Confirmar eliminación",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);

                    if (confirmacion == DialogResult.Yes)
                    {
                        bool resultado = await _excelService.EliminarGastoAsync(gasto);

                        if (resultado)
                        {
                            MessageBox.Show("Gasto eliminado correctamente", "Éxito",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                            // Recargar para actualizar totales
                            CargarDatos();
                        }
                        else
                        {
                            MessageBox.Show("No se pudo eliminar el gasto", "Error",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error al eliminar gasto: {ex.Message}", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void DgvGastos_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            // Pintar el botón eliminar en rojo (última columna, índice 6)
            if (e.ColumnIndex == 6 && e.RowIndex >= 0)
            {
                // Pintar el fondo de la celda primero
                e.PaintBackground(e.CellBounds, true);

                // Calcular un rectángulo más pequeño centrado en la celda
                var buttonSize = 24;
                var buttonRect = new Rectangle(
                    e.CellBounds.Left + (e.CellBounds.Width - buttonSize) / 2,
                    e.CellBounds.Top + (e.CellBounds.Height - buttonSize) / 2,
                    buttonSize,
                    buttonSize
                );

                // Dibujar el botón rojo con bordes redondeados
                using (var brush = new SolidBrush(Color.FromArgb(220, 38, 38))) // Rojo
                using (var path = new System.Drawing.Drawing2D.GraphicsPath())
                {
                    int radius = 4;
                    path.AddArc(buttonRect.X, buttonRect.Y, radius, radius, 180, 90);
                    path.AddArc(buttonRect.Right - radius, buttonRect.Y, radius, radius, 270, 90);
                    path.AddArc(buttonRect.Right - radius, buttonRect.Bottom - radius, radius, radius, 0, 90);
                    path.AddArc(buttonRect.X, buttonRect.Bottom - radius, radius, radius, 90, 90);
                    path.CloseFigure();
                    
                    e.Graphics.FillPath(brush, path);
                    
                    using (var pen = new Pen(Color.FromArgb(185, 28, 28), 1))
                    {
                        e.Graphics.DrawPath(pen, path);
                    }
                }

                // Dibujar el texto "✖" centrado
                var font = new Font(dgvGastos.Font.FontFamily, 11, FontStyle.Bold);
                var textSize = TextRenderer.MeasureText("✖", font);
                var textLocation = new Point(
                    buttonRect.Left + (buttonRect.Width - textSize.Width) / 2,
                    buttonRect.Top + (buttonRect.Height - textSize.Height) / 2
                );

                TextRenderer.DrawText(e.Graphics, "✖", font, textLocation, Color.White);

                e.Handled = true;
            }
        }
    }
}
