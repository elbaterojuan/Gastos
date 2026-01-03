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
        private Label lblTitulo;
        private Button btnAgregarGasto;
        private Button btnDashboard;
        private Button btnSalir;

        public FormMenuPrincipal()
        {
            InitializeComponent();
            InicializarServicios();
        }

        private void InitializeComponent()
        {
            this.Text = "Sistema de Gestión de Gastos";
            this.Size = new Size(500, 400);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = TemaColores.FondoClaro;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = true;

            // Título
            lblTitulo = new Label
            {
                Text = "💰 Sistema de Gestión de Gastos",
                Font = new Font("Segoe UI", 18F, FontStyle.Bold),
                ForeColor = TemaColores.TextoOscuro,
                AutoSize = false,
                Size = new Size(450, 50),
                Location = new Point(25, 20),
                TextAlign = ContentAlignment.MiddleCenter
            };

            // Botón Dashboard
            btnDashboard = new Button
            {
                Text = "📊 Ver Dashboard",
                Size = new Size(400, 60),
                Location = new Point(50, 90),
                Font = new Font("Segoe UI", 12F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnDashboard.AplicarEstiloBotonPrimario();
            btnDashboard.Click += BtnDashboard_Click;

            // Botón Agregar Gasto
            btnAgregarGasto = new Button
            {
                Text = "➕ Agregar Gasto",
                Size = new Size(400, 60),
                Location = new Point(50, 170),
                Font = new Font("Segoe UI", 12F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnAgregarGasto.BackColor = TemaColores.SecundarioVerde;
            btnAgregarGasto.ForeColor = Color.White;
            btnAgregarGasto.FlatStyle = FlatStyle.Flat;
            btnAgregarGasto.FlatAppearance.BorderSize = 0;
            btnAgregarGasto.Click += BtnAgregarGasto_Click;

            // Botón Salir
            btnSalir = new Button
            {
                Text = "🚪 Salir",
                Size = new Size(400, 50),
                Location = new Point(50, 250),
                Font = new Font("Segoe UI", 11F),
                Cursor = Cursors.Hand
            };
            btnSalir.AplicarEstiloBotonSecundario();
            btnSalir.Click += (s, e) => this.Close();

            // Agregar controles al formulario
            this.Controls.Add(lblTitulo);
            this.Controls.Add(btnDashboard);
            this.Controls.Add(btnAgregarGasto);
            this.Controls.Add(btnSalir);
        }

        private void InicializarServicios()
        {
            try
            {
                // Usar el directorio de la aplicación
                string directorioApp = AppDomain.CurrentDomain.BaseDirectory;
                string nombreArchivo = "Gastos.xlsm";
                
                _excelService = new ExcelService(directorioApp, nombreArchivo);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al inicializar:\n{ex.Message}\n\nArchivo: Gastos.xlsm\nUbicación: {AppDomain.CurrentDomain.BaseDirectory}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
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
