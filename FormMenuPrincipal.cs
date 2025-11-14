using System;
using System.Drawing;
using System.Windows.Forms;
using Gastos.Services;
using Gastos.Utils;
using Gastos.Views;

namespace Gastos
{
    public partial class FormMenuPrincipal : Form
    {
        private ExcelService _excelService;
        private Panel panelHeader;
        private Panel panelButtons;
        private Button btnAgregarGasto;
        private Button btnDashboard;
        private Button btnSalir;
        private Label lblTitulo;
        private Label lblSubtitulo;

        public FormMenuPrincipal()
        {
            InitializeComponent();
            InicializarServicios();
        }

        private void InitializeComponent()
        {
            this.Size = new Size(600, 500);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "Sistema de Gestión de Gastos";
            this.BackColor = TemaColores.FondoClaro;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;

            CrearHeader();
            CrearBotones();
        }

        private void InicializarServicios()
        {
            try
            {
                _excelService = new ExcelService(
                    Properties.Settings.Default.Carpeta,
                    Properties.Settings.Default.Archivo
                );
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al inicializar Excel:\n{ex.Message}\n\nVerifique la configuración en Settings.",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
        }

        private void CrearHeader()
        {
            panelHeader = new Panel
            {
                Dock = DockStyle.Top,
                Height = 150,
                BackColor = TemaColores.PrimarioAzul
            };

            lblTitulo = new Label
            {
                Text = "💰 Gestión de Gastos",
                Font = new Font("Segoe UI", 24F, FontStyle.Bold),
                ForeColor = Color.White,
                AutoSize = true,
                Location = new Point(150, 40)
            };

            lblSubtitulo = new Label
            {
                Text = "Sistema moderno de control de gastos personales",
                Font = new Font("Segoe UI", 11F),
                ForeColor = Color.FromArgb(200, 255, 255, 255),
                AutoSize = true,
                Location = new Point(140, 85)
            };

            panelHeader.Controls.AddRange(new Control[] { lblTitulo, lblSubtitulo });
            this.Controls.Add(panelHeader);
        }

        private void CrearBotones()
        {
            panelButtons = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(80, 40, 80, 40)
            };

            // Botón Dashboard
            btnDashboard = CrearBotonMenu(
                "📊 Dashboard",
                "Ver estadísticas y gráficos de gastos",
                TemaColores.PrimarioAzul,
                0
            );
            btnDashboard.Click += BtnDashboard_Click;

            // Botón Agregar Gasto
            btnAgregarGasto = CrearBotonMenu(
                "➕ Agregar Gasto",
                "Registrar un nuevo gasto en Excel",
                TemaColores.SecundarioVerde,
                100
            );
            btnAgregarGasto.Click += BtnAgregarGasto_Click;

            // Botón Salir
            btnSalir = CrearBotonMenu(
                "🚪 Salir",
                "Cerrar la aplicación",
                TemaColores.TextoGris,
                200
            );
            btnSalir.Click += (s, e) => Application.Exit();

            panelButtons.Controls.AddRange(new Control[] { btnDashboard, btnAgregarGasto, btnSalir });
            this.Controls.Add(panelButtons);
        }

        private Button CrearBotonMenu(string texto, string tooltip, Color color, int yOffset)
        {
            var btn = new Button
            {
                Text = texto,
                Size = new Size(420, 60),
                Location = new Point(10, yOffset),
                BackColor = color,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 13F, FontStyle.Bold),
                Cursor = Cursors.Hand,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(20, 0, 0, 0)
            };
            btn.FlatAppearance.BorderSize = 0;

            var tooltipControl = new ToolTip();
            tooltipControl.SetToolTip(btn, tooltip);

            var colorOriginal = color;
            var colorHover = Color.FromArgb(
                Math.Max(0, color.R - 30),
                Math.Max(0, color.G - 30),
                Math.Max(0, color.B - 30)
            );

            btn.MouseEnter += (s, e) => btn.BackColor = colorHover;
            btn.MouseLeave += (s, e) => btn.BackColor = colorOriginal;

            return btn;
        }

        private void BtnDashboard_Click(object sender, EventArgs e)
        {
            try
            {
                var dashboard = new FormDashboard(_excelService);
                dashboard.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al abrir dashboard: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnAgregarGasto_Click(object sender, EventArgs e)
        {
            try
            {
                var formAgregar = new Form1();
                formAgregar.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al abrir formulario: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            _excelService?.Dispose();
        }
    }
}
