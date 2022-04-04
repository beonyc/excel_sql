using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace excel_kurs_Sugrobov
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "nORTHWNDDataSet.Employees". При необходимости она может быть перемещена или удалена.
            this.employeesTableAdapter.Fill(this.nORTHWNDDataSet.Employees);

        }

        private void btnExport_Click(object sender, EventArgs e)
        {

            using (SaveFileDialog forfilter = new SaveFileDialog() { Filter = "Excel WorkBook|*.xlsx" })
            {
                if (forfilter.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        using (XLWorkbook workBook = new XLWorkbook())
                        {
                            workBook.Worksheets.Add(this.nORTHWNDDataSet.Employees.CopyToDataTable(), "employees");
                            workBook.SaveAs(forfilter.FileName);
                        }
                   
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    }
                }
            }

        }
    }
}
