using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Threading;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;

namespace Gastos
{
    public partial class Form1 : Form
    {
        Microsoft.Office.Interop.Excel.Application oXL;
        Workbooks oWBs;
        Workbook oWB;

        public Form1()
        {
            oXL = getExcel();
            oWBs = oXL.Workbooks;
            oXL.WindowState = XlWindowState.xlMaximized;
            oXL.Visible = true;
            oWB = getWorkbook();
            Thread.CurrentThread.CurrentCulture = new CultureInfo("es-ES");
            Thread.CurrentThread.CurrentUICulture = new CultureInfo("es-ES");
            InitializeComponent();
            dateTimePicker1.Value = DateTime.Now;
            string[] strArray = new string[Properties.Settings.Default.Categorias.Count];
            Properties.Settings.Default.Categorias.CopyTo(strArray, 0);
            comboBox1.Items.AddRange(strArray);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Worksheet oSheet = getSheet(oWB, dateTimePicker1.Value);
            oSheet.Select(Type.Missing);
            addValues(oSheet, getLastRow(oSheet));
            cleanInputs();
        }


        private int getLastRow(Worksheet oSheet)
        {
            for (int i = 1; i <= oSheet.UsedRange.Rows.Count; i++)
            {
                if (oSheet.Cells[i, 1].Value == null && oSheet.Cells[i, 2].Value == null && oSheet.Cells[i, 3].Value == null)
                {
                    return i;
                }
            }

            return oSheet.UsedRange.Rows.Count;
        }

        private void addValues(Worksheet oSheet,int row)
        {
            oSheet.Cells[row, 1].Value = dateTimePicker1.Value.ToString("M/d/yyyy");
            oSheet.Cells[row, 2].Value = comboBox1.Text;
            oSheet.Cells[row, 3].Value = numericUpDown1.Value;
            oSheet.Cells[row, 4].Value = comboBox2.Text;
            oSheet.Cells[row, 5].Value = (checkBox1.Checked) ? "SI" : "NO";
            oSheet.Cells[row, 8].Value = textBox1.Text;
        }

        private void cleanInputs()
        {
            comboBox1.Text = null;
            numericUpDown1.Value = 1;
            comboBox2.Text = null;
            checkBox1.Checked = true;
            textBox1.Text = null;
        }

        private Worksheet getSheet(Workbook oWB,DateTime date)
        {
            string month = date.ToString("MMMM", Thread.CurrentThread.CurrentCulture);
            month = char.ToUpper(month[0]) + month.Substring(1);
            return oWB.Worksheets[String.Format("{0}-{1}",month,date.ToString("yy"))];
        }

        private Microsoft.Office.Interop.Excel.Application getExcel()
        {
            try
            {
                return (Microsoft.Office.Interop.Excel.Application)Marshal.GetActiveObject("Excel.Application");
            }
            catch
            {
                return new Microsoft.Office.Interop.Excel.Application();
            }
            
        }

        private Workbook getWorkbook()
        {
            try
            {
                return oWBs.get_Item(Properties.Settings.Default.Archivo);
            }
            catch
            {
                return oWBs.Open(Path.Combine(Properties.Settings.Default.Carpeta,Properties.Settings.Default.Archivo));
            }
            
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (oXL != null)
            {
                oWBs.Close();
                oXL.Quit();

                Marshal.ReleaseComObject(oWBs);
                Marshal.ReleaseComObject(oXL);
            }
        }
    }
}
