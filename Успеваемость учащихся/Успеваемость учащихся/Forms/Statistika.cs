using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Успеваемость_учащихся.Forms
{
    public partial class Statistika : Form
    {
        public Statistika()
        {
            InitializeComponent();
        }

        private void Statistika_Load(object sender, EventArgs e)
        {
            LoadGrid();
        }
        private void LoadGrid()
            {
                dataGridView1.DataSource = Class.Database.Query(@"SELECT [id] as Номер
      ,[Название класса]
      ,[Кол учащихся] as 'Всего учащихся'
      ,[Кол отличников] as 'Закончили на 5'
      ,[Количество ударников] as 'Закончили на 4'
      ,[Количество троечников] as 'Закончили на 3'
      ,[Пропуски по уваж причине] as 'Уважительные пропуски'
      ,[Прропуски по неуваж причине] as 'Неуважительные пропуски'
      ,[Успеваемость]
  FROM [dbo].[Uspev2]");
            }

        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Add();
            Excel.Worksheet worksheet = workbook.Sheets[1];
            worksheet.Name = "Отчёт";
            worksheet.Cells[2, 3] = "Статистика классов";
            Excel.Range rng1 = worksheet.Range[worksheet.Cells[2, 3], worksheet.Cells[2, 3]];
            rng1.Cells.Font.Name = "Times New Roman";
            rng1.Cells.Font.Size = 24;
            rng1.Font.Bold = true;
            rng1.Cells.Font.Color = ColorTranslator.ToOle(Color.Green);
            worksheet.Cells[4, 1] = "Номер";
            worksheet.Columns[1].ColumnWidth = 20;
            worksheet.Cells[4, 2] = "Название класса";
            worksheet.Columns[2].ColumnWidth = 20;
            worksheet.Cells[4, 3] = "Количество учащихся";
            worksheet.Columns[3].ColumnWidth = 20;
            worksheet.Cells[4, 4] = "Количество отличников";
            worksheet.Cells[4].ColumnWidth = 22;
            worksheet.Cells[4, 5] = "Количество ударников";
            worksheet.Columns[5].ColumnWidth = 22;
            worksheet.Cells[4, 6] = "Количество троечников";
            worksheet.Columns[6].ColumnWidth = 22;
            worksheet.Cells[4, 7] = "Пропусков по уваж причине";
            worksheet.Columns[7].ColumnWidth = 22;
            worksheet.Cells[4, 8] = "Пропусков по неуваж причине";
            worksheet.Columns[8].ColumnWidth = 22;
            worksheet.Cells[4, 9] = "Успеваемость класса";
            worksheet.Columns[9].ColumnWidth = 20;
            Excel.Range rng2 = worksheet.Range[worksheet.Cells[4, 1], worksheet.Cells[4, 9]];
            rng2.Font.Bold = true;
            string SqlText = "Select * from Uspev2";
            SqlDataAdapter adapter = new SqlDataAdapter(SqlText, Class.Database.constr);
            DataTable table = new DataTable();
            adapter.Fill(table);
            int i = 5;
            foreach (DataRow row in table.Rows)
            {
                worksheet.Cells[i, 1] = row["id"];
                worksheet.Cells[i, 2] = row["Название класса"];
                worksheet.Cells[i, 3] = row["Кол учащихся"];
                worksheet.Cells[i, 4] = row["Кол отличников"];
                worksheet.Cells[i, 5] = row["Количество ударников"];
                worksheet.Cells[i, 6] = row["Количество троечников"];
                worksheet.Cells[i, 7] = row["Пропуски по уваж причине"];
                worksheet.Cells[i, 8] = row["Прропуски по неуваж причине"];
                worksheet.Cells[i, 9] = row["Успеваемость"];
                i++;
                Excel.Range rng3 = worksheet.Range[worksheet.Cells[4, 1], worksheet.Cells[i - 1, 9]];
                rng3.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle =
                Excel.XlLineStyle.xlContinuous;
                rng3.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).LineStyle =
                Excel.XlLineStyle.xlContinuous;
                rng3.Borders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous;
                rng3.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous;
                rng3.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;

            }
            excelApp.Visible = true;
            excelApp.UserControl = true;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Menu f = new Menu();
            f.Show();
            this.Hide();
        }

        private void обновитьСтатистикуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            UpdateGrid();
        }
        private void UpdateGrid(bool showMessage = true)
        {
            DataTable updated = new DataTable();

            if (Class.Database.ExecuteSqlCommand(@"SELECT [id] as Номер
      ,[Название класса]
      ,[Кол учащихся]
      ,[Кол отличников]
      ,[Количество ударников]
      ,[Количество троечников]
      ,[Пропуски по уваж причине]
      ,[Прропуски по неуваж причине]
      ,[Успеваемость]
  FROM [dbo].[Uspev2]", updated))
            {
                dataGridView1.DataSource = updated;
                dataGridView1.Columns[0].Visible = false;
                if (showMessage)
                    MessageBox.Show("Данные статистики обновлены.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            updated.Dispose();
        }
    }
}

