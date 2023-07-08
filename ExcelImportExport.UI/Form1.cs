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

namespace ExcelImportExport.UI
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
            
                      
        }

         



        private void btnExport_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog() { Filter= "Excel Workbook|*.xlsx|JPeg Image|*.jpg|Bitmap Image|*.bmp|Gif Image|*.gif" })
            {
                if (sfd.ShowDialog()==DialogResult.OK)
                {
                    try
                    {
                        using (XLWorkbook workbook = new XLWorkbook())
                        {
                            workbook.Worksheets.Add(this.appData3.Employees.CopyToDataTable(), "Employees");
                            workbook.SaveAs(sfd.FileName);
                        }
                        MessageBox.Show("Dosya basariyla yuklendi","Message",MessageBoxButtons.OK,MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Hata: "+ex.Message,"Message",MessageBoxButtons.OK,MessageBoxIcon.Error);
                    }
                }
            }
        }
    }
}
