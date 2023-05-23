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
    public partial class Prepod : Form
    {
        public Prepod()
        {
            InitializeComponent();
        }

        private void Prepod_Load(object sender, EventArgs e)
        {
            LoadGrid();
        }
        private void LoadGrid()
        {
            dataGridView1.DataSource = Class.Database.Query(@"SELECT [Код_преподавателя] as Номер
                                                                                       ,[Фамилия]
                                                                                       ,[Имя]
                                                                                       ,[Отчество]
                                                                                       ,[Стаж] as 'Стаж работы'
                                                                                       ,[Предмет]
                                                                                       FROM [dbo].[prepod]");

        }

        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = Class.Database.Query($@"INSERT INTO [dbo].[prepod]
                                                                     ([Код_преподавателя]
                                                                     ,[Фамилия]
                                                                     ,[Имя]
                                                                     ,[Отчество]
                                                                     ,[Стаж]
                                                                     ,[Предмет])
                                                               VALUES
                                                                     ('{textBox1.Text}',
                                                                      '{textBox2.Text}',
                                                                      '{textBox3.Text}',
                                                                      '{textBox4.Text}',
                                                                      '{textBox5.Text}',
                                                                      '{textBox6.Text}')"); LoadGrid();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = Class.Database.Query($@"UPDATE [dbo].[prepod]
                                                            SET 
                                                            [Код_преподавателя] =   '{textBox1.Text}'
                                                            ,[Фамилия] =            '{textBox2.Text}'
                                                            ,[Имя] =                '{textBox3.Text}'
                                                            ,[Отчество] =           '{textBox4.Text}'
                                                            ,[Стаж] =               '{textBox5.Text}'
                                                            ,[Предмет] =            '{textBox6.Text}'
                                                            WHERE [Код_преподавателя] = {dataGridView1.CurrentRow.Cells[0].Value}"); LoadGrid();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = Class.Database.Query($@"DELETE FROM [dbo].[prepod]
      WHERE [Код_преподавателя] = {dataGridView1.CurrentRow.Cells[0].Value}"); LoadGrid();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Add();
            Excel.Worksheet worksheet = workbook.Sheets[1];
            worksheet.Name = "Отчёт";
            worksheet.Cells[2, 3] = "Список преподавателей";
            Excel.Range rng1 = worksheet.Range[worksheet.Cells[2, 3], worksheet.Cells[2, 3]];
            rng1.Cells.Font.Name = "Times New Roman";
            rng1.Cells.Font.Size = 24;
            rng1.Font.Bold = true;
            rng1.Cells.Font.Color = ColorTranslator.ToOle(Color.Green);
            worksheet.Cells[4, 1] = "Номер";
            worksheet.Columns[1].ColumnWidth = 20;
            worksheet.Cells[4, 2] = "Фамилия";
            worksheet.Columns[2].ColumnWidth = 18;
            worksheet.Cells[4, 3] = "Имя";
            worksheet.Columns[3].ColumnWidth = 18;
            worksheet.Cells[4, 4] = "Отчество";
            worksheet.Cells[4].ColumnWidth = 18;
            worksheet.Cells[4, 5] = "Стаж";
            worksheet.Columns[5].ColumnWidth = 15;
            worksheet.Cells[4, 6] = "Предмет";
            worksheet.Columns[6].ColumnWidth = 15;
            Excel.Range rng2 = worksheet.Range[worksheet.Cells[4, 1], worksheet.Cells[4, 9]];
            rng2.Font.Bold = true;
            string SqlText = "Select * from prepod";
            SqlDataAdapter adapter = new SqlDataAdapter(SqlText, Class.Database.constr);
            DataTable table = new DataTable();
            adapter.Fill(table);
            int i = 5;
            foreach (DataRow row in table.Rows)
            {
                worksheet.Cells[i, 1] = row["Код_преподавателя"];
                worksheet.Cells[i, 2] = row["Фамилия"];
                worksheet.Cells[i, 3] = row["Имя"];
                worksheet.Cells[i, 4] = row["Отчество"];
                worksheet.Cells[i, 5] = row["Стаж"];
                worksheet.Cells[i, 6] = row["Предмет"];
                i++;
                Excel.Range rng3 = worksheet.Range[worksheet.Cells[4, 1], worksheet.Cells[i - 1, 6]];
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

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = Class.Database.Query(@"SELECT [Код_преподавателя] as Номер 
                                                        ,[Фамилия] 
                                                        ,[Имя] 
                                                        ,[Отчество] 
                                                        ,[Стаж] 
                                                        ,[Предмет]
                                                  FROM [dbo].[prepod] WHERE [Фамилия]  LIKE '" + textBox7.Text + "%' or [Имя] LIKE '" + textBox7.Text + "%' or [Отчество] LIKE '" + textBox7.Text + "%' or [Стаж] LIKE '" + textBox7.Text + "%' or [Код_преподавателя] LIKE '" + textBox7.Text + "%'");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Menu f = new Menu();
            f.Show();
            this.Hide();
        }

        private void обновитьДанныеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            UpdateGrid();
        }
        private void UpdateGrid(bool showMessage = true)
        {
            DataTable updated = new DataTable();

            if (Class.Database.ExecuteSqlCommand(@"SELECT [Код_преподавателя]  as Номер
      ,[Фамилия]
      ,[Имя]
      ,[Отчество]
      ,[Стаж]
      ,[Предмет]
  FROM [dbo].[prepod]", updated))
            {
                dataGridView1.DataSource = updated;
                dataGridView1.Columns[0].Visible = false;
                if (showMessage)
                    MessageBox.Show("Данные обновлены.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            updated.Dispose();
        }
    }
}

