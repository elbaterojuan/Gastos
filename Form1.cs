using System;
using System.Drawing;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Gastos.Models;
using Gastos.Services;
using Gastos.Utils;

namespace Gastos
{
    public partial class Form1 : Form
    {
        private ExcelService _excelService;
        private System.Windows.Forms.Timer _timer;

        public Form1()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("es-ES");
            Thread.CurrentThread.CurrentUICulture = new CultureInfo("es-ES");
            
            InitializeComponent();
            InicializarServicios();
            AplicarEstilosModernos();
            ConfigurarControles();
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
                MessageBox.Show($"Error al inicializar Excel: {ex.Message}", 
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
        }

        private void AplicarEstilosModernos()
        {
            // Estilo del formulario
            this.BackColor = TemaColores.FondoClaro;
            this.Font = new Font("Segoe UI", 9.5F);
            this.Size = new Size(500, 550);
            this.StartPosition = FormStartPosition.CenterScreen;
            
            // Estilos de labels
            label1.AplicarEstiloLabel();
            label2.AplicarEstiloLabel();
            label3.AplicarEstiloLabel();
            label4.AplicarEstiloLabel();
            label5.AplicarEstiloLabel();
            label6.AplicarEstiloLabel();
            label7.AplicarEstiloLabel();
            
            // Estilos de controles
            textBox1.AplicarEstiloTextBox();
            comboBox1.AplicarEstiloComboBox();
            comboBox2.AplicarEstiloComboBox();
            button1.AplicarEstiloBotonPrimario();
            button1.Text = "💾 Agregar Gasto";
            
            // Estilo del mensaje de éxito
            resultLbl.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            resultLbl.ForeColor = TemaColores.Exito;
            
            // Labels con iconos
            label1.Text = "📅 Fecha";
            label2.Text = "🏷️ Categoría";
            label3.Text = "💰 Monto";
            label4.Text = "👤 Quién pagó?";
            label5.Text = "⚖️ Gasto Proporcional?";
            label6.Text = "📝 Comentarios";
            label7.Text = "🔢 Cuotas";
        }

        private void ConfigurarControles()
        {
            dateTimePicker1.Value = DateTime.Now;
            
            // Cargar categorías
            string[] categorias = new string[Properties.Settings.Default.Categorias.Count];
            Properties.Settings.Default.Categorias.CopyTo(categorias, 0);
            comboBox1.Items.AddRange(categorias);
            
            // Configurar timer para ocultar mensaje de éxito
            _timer = new System.Windows.Forms.Timer { Interval = 3000 };
            _timer.Tick += Timer_Tick;
            
            // Tooltips informativos
            var tooltip = new ToolTip();
            tooltip.SetToolTip(checkBox1, "Si está marcado, el gasto se dividirá proporcionalmente");
            tooltip.SetToolTip(cuotas, "Número de cuotas para pagos diferidos");
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            if (!ValidarCampos()) return;

            try
            {
                button1.Enabled = false;
                button1.Text = "⏳ Guardando...";

                var gasto = new Gasto
                {
                    Fecha = dateTimePicker1.Value,
                    Categoria = comboBox1.Text,
                    Monto = numericUpDown1.Value,
                    QuienPago = comboBox2.Text,
                    EsProporcional = checkBox1.Checked,
                    Cuotas = (int)cuotas.Value,
                    Comentarios = textBox1.Text
                };

                await _excelService.AgregarGastoAsync(gasto);

                MostrarMensajeExito();
                LimpiarCampos();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al guardar el gasto: {ex.Message}", 
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                button1.Enabled = true;
                button1.Text = "💾 Agregar Gasto";
            }
        }

        private bool ValidarCampos()
        {
            if (string.IsNullOrWhiteSpace(comboBox1.Text))
            {
                MessageBox.Show("Por favor seleccione una categoría", "Validación", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                comboBox1.Focus();
                return false;
            }

            if (numericUpDown1.Value <= 0)
            {
                MessageBox.Show("El monto debe ser mayor a cero", "Validación", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                numericUpDown1.Focus();
                return false;
            }

            if (string.IsNullOrWhiteSpace(comboBox2.Text))
            {
                MessageBox.Show("Por favor indique quién pagó", "Validación", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                comboBox2.Focus();
                return false;
            }

            return true;
        }

        private void MostrarMensajeExito()
        {
            resultLbl.Text = "✓ Gasto agregado correctamente";
            resultLbl.Visible = true;
            _timer.Start();
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            resultLbl.Visible = false;
            _timer.Stop();
        }

        private void LimpiarCampos()
        {
            comboBox1.Text = "";
            numericUpDown1.Value = 1;
            comboBox2.Text = "";
            checkBox1.Checked = true;
            cuotas.Value = 1;
            textBox1.Text = "";
            comboBox1.Focus();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            _excelService?.Dispose();
            _timer?.Dispose();
        }
    }
}
