using System;
using System.Configuration;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
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
                // Usar el directorio de la aplicación
                string directorioApp = AppDomain.CurrentDomain.BaseDirectory;
                string nombreArchivo = "Gastos.xlsm";
                
                _excelService = new ExcelService(directorioApp, nombreArchivo);
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
            label4.Text = "👤 Quién Pagó";
            label7.Text = "🔢 Cuotas";
            label5.Text = "⚖️ ¿Gasto Proporcional?";
            label6.Text = "📝 Comentarios";
        }

        private void ConfigurarControles()
        {
            dateTimePicker1.Value = DateTime.Now;
            
            // Configurar formato de moneda
            var culture = new CultureInfo("es-AR"); // Argentina usa $
            culture.NumberFormat.CurrencySymbol = "$";
            numericUpDown1.Text = culture.NumberFormat.CurrencySymbol;
            
            // Configurar cuotas con valor por defecto 1
            cuotas.Value = 1;
            cuotas.Minimum = 1;
            cuotas.Maximum = 36;
            
            // Cargar categorías directamente desde el archivo .exe.config
            comboBox1.Items.Clear();
            
            try
            {
                var configPath = AppDomain.CurrentDomain.SetupInformation.ConfigurationFile;
                var configDoc = XDocument.Load(configPath);
                
                var categoriasElement = configDoc.Descendants("setting")
                    .FirstOrDefault(s => s.Attribute("name")?.Value == "Categorias");
                
                if (categoriasElement != null)
                {
                    var arrayOfString = categoriasElement.Descendants("string");
                    foreach (var categoria in arrayOfString)
                    {
                        comboBox1.Items.Add(categoria.Value);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al cargar categorías: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            
            // Configurar checkbox con valor por defecto
            checkBox1.Checked = true;
            
            // Configurar timer para ocultar mensaje de éxito
            _timer = new System.Windows.Forms.Timer { Interval = 3000 };
            _timer.Tick += Timer_Tick;
            
            // Tooltips informativos
            var tooltip = new ToolTip();
            tooltip.SetToolTip(checkBox1, "Marcar si el gasto debe dividirse proporcionalmente");
            tooltip.SetToolTip(textBox1, "Comentarios adicionales sobre el gasto");
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            if (!ValidarCampos()) return;

            try
            {
                button1.Enabled = false;
                button1.Text = "⏳ Guardando...";

                int cantidadCuotas = (int)cuotas.Value;
                decimal montoPorCuota = numericUpDown1.Value / cantidadCuotas;
                
                // Procesar cada cuota
                for (int i = 1; i <= cantidadCuotas; i++)
                {
                    var fechaCuota = dateTimePicker1.Value.AddMonths(i - 1);
                    
                    // Solo agregar info de cuota si hay más de una
                    string comentarioCuota;
                    if (cantidadCuotas > 1)
                    {
                        comentarioCuota = string.IsNullOrWhiteSpace(textBox1.Text) 
                            ? $"Cuota {i}/{cantidadCuotas}"
                            : $"{textBox1.Text} - Cuota {i}/{cantidadCuotas}";
                    }
                    else
                    {
                        comentarioCuota = textBox1.Text;
                    }
                    
                    var gasto = new Gasto
                    {
                        Fecha = fechaCuota,
                        Categoria = comboBox1.Text,
                        Monto = montoPorCuota,
                        QuienPago = comboBox2.Text,
                        EsProporcional = checkBox1.Checked,
                        Comentarios = comentarioCuota,
                        CantidadCuotas = cantidadCuotas
                    };

                    await _excelService.AgregarGastoAsync(gasto);
                }

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
            cuotas.Value = 1;
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
