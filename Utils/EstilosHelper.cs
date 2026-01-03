using System.Drawing;
using System.Windows.Forms;

namespace Gastos.Utils
{
    /// <summary>
    /// Extensiones para aplicar estilos modernos a controles
    /// </summary>
    public static class EstilosHelper
    {
        public static void AplicarEstiloBotonPrimario(this Button boton)
        {
            boton.BackColor = TemaColores.PrimarioAzul;
            boton.ForeColor = Color.White;
            boton.FlatStyle = FlatStyle.Flat;
            boton.FlatAppearance.BorderSize = 0;
            boton.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            boton.Cursor = Cursors.Hand;
            boton.Padding = new Padding(15, 8, 15, 8);
            boton.Height = 40;

            boton.MouseEnter += (s, e) => boton.BackColor = TemaColores.PrimarioAzulOscuro;
            boton.MouseLeave += (s, e) => boton.BackColor = TemaColores.PrimarioAzul;
        }

        public static void AplicarEstiloBotonSecundario(this Button boton)
        {
            boton.BackColor = TemaColores.FondoGris;
            boton.ForeColor = TemaColores.TextoOscuro;
            boton.FlatStyle = FlatStyle.Flat;
            boton.FlatAppearance.BorderSize = 1;
            boton.FlatAppearance.BorderColor = TemaColores.BordeGris;
            boton.Font = new Font("Segoe UI", 9.5F);
            boton.Cursor = Cursors.Hand;
            boton.Padding = new Padding(15, 8, 15, 8);
            boton.Height = 40;

            boton.MouseEnter += (s, e) => boton.BackColor = Color.FromArgb(230, 230, 230);
            boton.MouseLeave += (s, e) => boton.BackColor = TemaColores.FondoGris;
        }

        public static void AplicarEstiloLabel(this Label label, bool esHeader = false)
        {
            if (esHeader)
            {
                label.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
                label.ForeColor = TemaColores.TextoOscuro;
            }
            else
            {
                label.Font = new Font("Segoe UI", 9.5F);
                label.ForeColor = TemaColores.TextoGris;
            }
        }

        public static void AplicarEstiloTextBox(this System.Windows.Forms.TextBox textBox)
        {
            textBox.Font = new Font("Segoe UI", 10F);
            textBox.BorderStyle = BorderStyle.FixedSingle;
            textBox.Height = 32;
        }

        public static void AplicarEstiloComboBox(this ComboBox comboBox)
        {
            comboBox.Font = new Font("Segoe UI", 10F);
            comboBox.FlatStyle = FlatStyle.Flat;
            comboBox.Height = 32;
        }

        public static void AplicarEstiloPanel(this Panel panel, bool conSombra = false)
        {
            panel.BackColor = TemaColores.FondoBlanco;
            panel.Padding = new Padding(20);

            if (conSombra)
            {
                panel.Paint += (s, e) =>
                {
                    ControlPaint.DrawBorder(e.Graphics, panel.ClientRectangle,
                        TemaColores.BordeGris, 1, ButtonBorderStyle.Solid,
                        TemaColores.BordeGris, 1, ButtonBorderStyle.Solid,
                        TemaColores.BordeGris, 1, ButtonBorderStyle.Solid,
                        TemaColores.BordeGris, 1, ButtonBorderStyle.Solid);
                };
            }
        }

        public static void AplicarEstiloDataGridView(this DataGridView dgv)
        {
            dgv.BackgroundColor = TemaColores.FondoBlanco;
            dgv.BorderStyle = BorderStyle.None;
            dgv.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            dgv.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
            dgv.EnableHeadersVisualStyles = false;
            
            dgv.ColumnHeadersDefaultCellStyle.BackColor = TemaColores.PrimarioAzul;
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            dgv.ColumnHeadersDefaultCellStyle.Padding = new Padding(10);
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dgv.ColumnHeadersHeight = 45;
            dgv.ColumnHeadersVisible = true;
            
            dgv.DefaultCellStyle.Font = new Font("Segoe UI", 9.5F);
            dgv.DefaultCellStyle.BackColor = Color.White;
            dgv.DefaultCellStyle.ForeColor = TemaColores.TextoOscuro;
            dgv.DefaultCellStyle.SelectionBackColor = TemaColores.PrimarioAzul;
            dgv.DefaultCellStyle.SelectionForeColor = Color.White;
            dgv.DefaultCellStyle.Padding = new Padding(8);
            
            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 248, 250);
            dgv.RowTemplate.Height = 40;
            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgv.MultiSelect = false;
            dgv.AllowUserToAddRows = false;
            dgv.AllowUserToDeleteRows = false;
            dgv.ReadOnly = true;
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }
    }
}
