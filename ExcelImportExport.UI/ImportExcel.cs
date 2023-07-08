using DocumentFormat.OpenXml.Bibliography;
using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Metadata.W3cXsd2001;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Z.Dapper.Plus;

namespace ExcelImportExport.UI
{
    public partial class ImportExcel : Form
    {
        public ImportExcel()
        {
            InitializeComponent();
        }

        private void cmbTablo_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt = tableCollection[cmbTablo.SelectedItem.ToString()];
            //dataGridView1.DataSource = dt;
            if (dt != null)
            {
                List<Class1> employeeList = new List<Class1>();
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    Class1 employee = new Class1();
                    employee.EmployeeID = int.Parse(dt.Rows[i]["EmployeeID"].ToString());
                    employee.FirstName = dt.Rows[i]["FirstName"].ToString();
                    employee.LastName = dt.Rows[i]["LastName"].ToString();
                    employee.Title = dt.Rows[i]["Title"].ToString();
                    employee.TitleOfCourtesy = dt.Rows[i]["TitleOfCourtesy"].ToString();
                    employee.Address = dt.Rows[i]["Address"].ToString();
                    employee.City = dt.Rows[i]["City"].ToString();
                    employee.Region = dt.Rows[i]["Region"].ToString();
                    employee.PostalCode = dt.Rows[i]["PostalCode"].ToString();
                    employee.Country = dt.Rows[i]["Country"].ToString();
                    employee.HomePhone = dt.Rows[i]["HomePhone"].ToString();
                    employee.Extension = dt.Rows[i]["Extension"].ToString();
                    employee.Notes = dt.Rows[i]["Notes"].ToString();
                    
                    // todo string datetime bakalım
                    employeeList.Add(employee);

                }
                employeesBindingSource.DataSource = employeeList;
            }
            

        }
        DataTableCollection tableCollection;
        private void btnDLL_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx" })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtDosyaKonum.Text = openFileDialog.FileName;
                    using (var stream = File.Open(openFileDialog.FileName, FileMode.Open, FileAccess.Read))
                    {
                        using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {

                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }

                            });
                            tableCollection = result.Tables;
                            cmbTablo.Items.Clear();
                            foreach (DataTable table in tableCollection)
                            {

                                cmbTablo.Items.Add(table.TableName);// tabloalara ekliyor.
                            }
                        }
                    }
                }
            }
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            try
            {
                DapperPlusManager.Entity<Class1>().Table("Employees");
                List<Class1> employee = employeesBindingSource.DataSource as List<Class1>;
                if (employee!= null)
                {
                    using (IDbConnection db = new SqlConnection("server =DESKTOP-IH2SMGI; Database = Northwnd; trusted_connection = true"))
                    {
                        db.BulkInsert(employee);
                    }
                }
                MessageBox.Show("İşlem tamamlandı.");
                
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ImportExcel_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'appDataNorth.Employees' table. You can move, or remove it, as needed.
            this.employeesTableAdapter.Fill(this.appDataNorth.Employees);
            // TODO: This line of code loads data into the 'nORTHWNDDataSet1.Employees' table. You can move, or remove it, as needed.
            //this.employeesTableAdapter.Fill(this.nORTHWNDDataSet1.Employees);
            // TODO: This line of code loads data into the 'nORTHWNDDataSet.Employees' table. You can move, or remove it, as needed.
            //this.employeesTableAdapter.Fill(this.nORTHWNDDataSet.Employees);

        }

        
    }
}
