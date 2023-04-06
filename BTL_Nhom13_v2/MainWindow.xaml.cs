using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data.SqlClient;
using System.Security.Cryptography;
using DevExpress.Utils.CommonDialogs.Internal;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System.Data;
using System.IO;
using System.Windows.Forms;
using Path = System.IO.Path;
using System.ComponentModel;

namespace BTL_Nhom13_v2
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            dlg.DefaultExt = ".xml";


            Nullable<bool> result = dlg.ShowDialog();
            if (result == true)
            {
                string Filename = dlg.FileName;
                txtBloxOpenFile.Text = Filename;
            }

        }

        private void btnConvertExcel_Click(object sender, RoutedEventArgs e)
        {
            progressar1.Value = 0;
            if (txtBloxOpenFile.Text != "" && txtBoxSaveFile.Text != "")
            {
                if (File.Exists(txtBloxOpenFile.Text))
                {
                    string CustXmlFilePath = Path.Combine(new FileInfo(txtBloxOpenFile.Text).DirectoryName, txtBoxSaveFile.Text);
                    System.Data.DataTable dt = CreateDataTableFromXml(txtBloxOpenFile.Text);
                    //LiceLicenseContext = LicenseContext.NonCommercial;
                    ExportDataTableToExcel(dt, CustXmlFilePath);

                    MessageBox.Show("Hoan thanh");
                }
            }
            else if (txtBloxOpenFile.Text != "")
            {
                if(File.Exists(txtBloxOpenFile.Text))
                {
                    FileInfo fi = new FileInfo(txtBloxOpenFile.Text);
                    string XlFile = fi.DirectoryName + "\\" + fi.Name.Replace(fi.Extension, ".xlsx");
                    System.Data.DataTable dt = CreateDataTableFromXml(txtBloxOpenFile.Text);
                    ExportDataTableToExcel(dt, XlFile);

                    MessageBox.Show("Hoan thanh 2");

                }
            }
            else
            {
                MessageBox.Show("Nhap lai");
            }
        }

        private void ExportDataTableToExcel(System.Data.DataTable table, string Xlfile)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook book = excel.Application.Workbooks.Add(Type.Missing);
            excel.Visible = false;
            excel.DisplayAlerts = false;
            Worksheet excelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)book.ActiveSheet;
            excelWorkSheet.Name = table.TableName;

            progressar1.Maximum = table.Columns.Count; 
            for(int i = 1; i < table.Columns.Count + 1; i++)
            {
                excelWorkSheet.Cells[1, i] = table.Columns[i - 1].ColumnName;
                if(progressar1.Value < progressar1.Maximum)
                {
                    progressar1.Value++;
                    int percent = (int)(((double)progressar1.Value / (double)progressar1.Maximum) * 100);
                    //object value = progressar1.CreateGraphics().DrawString(percent.ToString() + "%", new System.Drawing.Font("Arial", (float)8.25, FontStyle.Regular), Brushes.Black, new PointF(progressar1.Width / 2 - 10, progressar1.Height / 2 - 7));
                    object value = progressar1.Value;
                    //object value1 = System.Windows.Application.DoEvents();
                }
            }
            for (int j = 0; j < table.Rows.Count  ; j++)
            {
                for(int k = 0; k < table.Rows.Count; k++)
                {
                    excelWorkSheet.Cells[j + 2 , k+1 ] = table.Rows[j].ItemArray[k].ToString();
                    
                }

                
            }

            book.SaveAs(Xlfile);
            book.Close(true);
            excel.Quit();

            Marshal.ReleaseComObject(book);
            Marshal.ReleaseComObject(book);
            Marshal.ReleaseComObject(excel);

        }

        public System.Data.DataTable CreateDataTableFromXml(string XmlFile)
        {
            System.Data.DataTable Dt = new System.Data.DataTable();
            try
            {
                DataSet ds = new DataSet();
                ds.ReadXml(XmlFile);
                Dt.Load(ds.CreateDataReader());
            }catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return Dt;
        }
    }
}
