using System;
using System.Data;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
namespace WinFormsApp1
{
    public partial class Form1 : Form
    {
        private string file = "";

        public Form1()
        {
            InitializeComponent();
        }

        private void OpenFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xlsx;*.xls";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                file = openFileDialog.FileName;
                ZagruzitDataGrid(file);
            }
        }

        private void ZagruzitDataGrid(string put)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook rabKniga = excelApp.Workbooks.Open(put, ReadOnly: false);
            Excel._Worksheet list = (Excel._Worksheet)rabKniga.Sheets[1];
            Excel.Range diapazon = list.UsedRange;

            DataTable table = new DataTable();

            for (int j = 1; j <= diapazon.Columns.Count; j++)
            {
                table.Columns.Add("Kolonka " + j);
            }

            for (int i = 1; i <= diapazon.Rows.Count; i++)
            {
                DataRow stroka = table.NewRow();
                for (int j = 1; j <= diapazon.Columns.Count; j++)
                {
                    stroka[j - 1] = (diapazon.Cells[i, j] as Excel.Range).Text;
                }
                table.Rows.Add(stroka);
            }

            dataGridView1.DataSource = table;

            rabKniga.Close(false);
            excelApp.Quit();
        }

        private void Change_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(file))
            {
                MessageBox.Show("файл не выбран(");
                return;
            }

            if (string.IsNullOrEmpty(textBoxYacheika.Text))
            {
                MessageBox.Show("адрес €чейки пустой");
                return;
            }
            try
            {
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook rabKniga = excelApp.Workbooks.Open(file, ReadOnly: false);
                Excel._Worksheet list = (Excel._Worksheet)rabKniga.Sheets[1];

                string yacheika = textBoxYacheika.Text;
                string tekst = textBoxText.Text;

                Excel.Range range = list.Range[yacheika];
                range.Value2 = tekst;

                rabKniga.Save();
                MessageBox.Show("данные сохранились");

                ZagruzitDataGrid(file);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ќшибка {ex.Message}");
            }
        }
    }   
}
