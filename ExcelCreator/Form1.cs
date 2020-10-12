using System;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace ExcelCreator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
        }

        private void selectFilesButton_Click(object sender, CancelEventArgs e)
        {
            label1.Visible = false;

            try
            {
                var excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;
                var workBook = excel.Workbooks.Add(Type.Missing);
                var workSheet = (Worksheet)workBook.ActiveSheet;
                workSheet.Name = "Samples";
                dynamic celLrangE = null;

                string[] fileNames = this.openFileDialog1.SafeFileNames;
                var table = new System.Data.DataTable();

                if (fileNames != null && fileNames.Any())
                {
                    table.Columns.Add("Sample", typeof(string));
                    table.Columns.Add("Fat", typeof(decimal));
                    table.Columns.Add("SNF", typeof(decimal));
                    table.Columns.Add("Protein", typeof(decimal));
                    table.Columns.Add("Device", typeof(string));
                    table.Columns.Add("Temperature", typeof(decimal));
                    table.Columns.Add("Clr", typeof(decimal));
                    table.Columns.Add("Date", typeof(string));

                    foreach (string fileName in fileNames)
                    {
                        var splittedValues = fileName.Split('_');

                        string dataString = splittedValues[1];
                        var indexOfFat = dataString.IndexOf('F');
                        string dataStringWithoutBuffalo = dataString.Substring(indexOfFat);
                        dataStringWithoutBuffalo = dataString.Substring(indexOfFat).Replace("F", "*")
                            .Replace("S", "*")
                            .Replace("P", "*");
                        var milkContents = dataStringWithoutBuffalo.Split('*');
                        var indexOfSnf = dataStringWithoutBuffalo.IndexOf('S');
                        var indexOfProtein = dataStringWithoutBuffalo.IndexOf('P');
                        decimal fat = Convert.ToDecimal(milkContents[1]);
                        decimal snf = Convert.ToDecimal(milkContents[2]);
                        decimal protein = Convert.ToDecimal(milkContents[3]);

                        string device = splittedValues[2];
                        decimal temperature = Convert.ToDecimal(splittedValues[3].Substring(1));
                        string date = splittedValues[4].Split('.')[0];

                        var clr = (snf - (fat * 0.20m) - 0.66m) / 0.25m;

                        this.AddDataToTable(ref table, fileName.Substring(0, fileName.Length - 4), fat, snf, protein, device, temperature, clr, date);
                    }
                }

                int rowcount = 2;

                foreach (DataRow datarow in table.Rows)
                {
                    rowcount += 1;
                    for (int i = 1; i <= table.Columns.Count; i++)
                    {

                        if (rowcount == 3)
                        {
                            workSheet.Cells[2, i] = table.Columns[i - 1].ColumnName;
                            workSheet.Cells.Font.Color = Color.Black;
                        }

                        workSheet.Cells[rowcount, i] = datarow[i - 1].ToString();

                        if (rowcount > 3)
                        {
                            if (i == table.Columns.Count)
                            {
                                if (rowcount % 2 == 0)
                                {
                                    celLrangE = workSheet.Range[workSheet.Cells[rowcount, 1], workSheet.Cells[rowcount, table.Columns.Count]];
                                }
                            }
                        }
                    }
                }

                celLrangE = workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[rowcount, table.Columns.Count]];
                celLrangE.EntireColumn.AutoFit();
                Borders border = celLrangE.Borders;
                border.LineStyle = XlLineStyle.xlContinuous;
                border.Weight = 2d;
                celLrangE = workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[2, table.Columns.Count]];
                var currentDate = DateTime.Now;
                var excelFileName = $"{Path.GetDirectoryName(this.openFileDialog1.FileName)}\\Samples_{currentDate.Year}-{currentDate.Month}-{currentDate.Day}-{currentDate.Hour}-{currentDate.Minute}-{currentDate.Second}.xls";
                workBook.SaveAs(excelFileName);
                workBook.Close();
                excel.Quit();

                label1.Text = $"Successfully Created Excel File at {excelFileName}";
                label1.ForeColor = Color.Blue;
                label1.Visible = true;
            }
            catch (Exception ex)
            {
                label1.Text = $"Some Error Occured: {ex.Message}";
                label1.ForeColor = Color.Red;
                label1.Visible = true;
            }
        }

        public void AddDataToTable(ref System.Data.DataTable table, string sample, decimal fat, decimal snf, decimal protein, string device,
            decimal temperature, decimal clr, string date)
        {
            table.Rows.Add(sample, fat, snf, protein, device, temperature, clr, date);
        }
    }
}
