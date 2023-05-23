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
    public partial class Uspevaemost : Form
    {
        public Uspevaemost()
        {
            InitializeComponent();
        }

        private void Uspevaemost_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "uspevaemotsDataSet.Class1". При необходимости она может быть перемещена или удалена.
            this.class1TableAdapter.Fill(this.uspevaemotsDataSet.Class1);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "uspevaemotsDataSet.Ucheniki". При необходимости она может быть перемещена или удалена.
            this.uchenikiTableAdapter.Fill(this.uspevaemotsDataSet.Ucheniki);
            LoadGrid();
        }
        private void LoadGrid()
        {
            dataGridView1.DataSource = Class.Database.Query(@"SELECT [id] as 'Номер п/п'
      ,[ФИО]
      ,[Класс]
      ,[Пропуски по уваж]
      ,[Пропуски по неуваж]
      ,[Математика]
      ,[Русский язык]
      ,[Английский язык]
      ,[Литература]
      ,[Физкультура]
      ,[Физика]
      ,[Сред.балл]
  FROM [dbo].[uspevaemost f]");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = Class.Database.Query($@"INSERT INTO [dbo].[uspevaemost f]
                                                                                       ([id]
                                                                                       ,[ФИО]
                                                                                       ,[Класс]
                                                                                       ,[Пропуски по уваж]
                                                                                       ,[Пропуски по неуваж]
                                                                                       ,[Математика]
                                                                                       ,[Русский язык]
                                                                                       ,[Английский язык]
                                                                                       ,[Литература]
                                                                                       ,[Физкультура]
                                                                                       ,[Физика]
                                                                                       ,[Сред.балл])
                                                                                 VALUES
                                                                                    ('{comboBox3.Text}',
                                                                                    '{comboBox2.Text}',
                                                                                    '{comboBox1.Text}',
                                                                                    '{textBox3.Text}',
                                                                                    '{textBox4.Text}',
                                                                                    '{textBox5.Text}',
                                                                                    '{textBox6.Text}',
                                                                                    '{textBox7.Text}',
                                                                                    '{textBox8.Text}',
                                                                                    '{textBox9.Text}',
                                                                                    '{textBox10.Text}',
                                                                                    '{textBox11.Text}')"); LoadGrid();
                                                                                    MessageBox.Show("Запись успешно добавлена", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                                                                }

        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = Class.Database.Query($@"UPDATE [dbo].[uspevaemost f]
                                                               SET 
                                                                   [id] = '{comboBox3.Text}'
                                                                  ,[ФИО] =   '{comboBox2.Text}'
                                                                  ,[Класс] =  '{comboBox1.Text}'
                                                                  ,[Пропуски по уваж] =     '{textBox3.Text}'
                                                                  ,[Пропуски по неуваж] =   '{textBox4.Text}'
                                                                  ,[Математика] =      '{textBox5.Text}'
                                                                  ,[Русский язык] =    '{textBox6.Text}'
                                                                  ,[Английский язык] = '{textBox7.Text}'
                                                                  ,[Литература] =  '{textBox8.Text}'
                                                                  ,[Физкультура] = '{textBox9.Text}'
                                                                  ,[Физика] =      '{textBox10.Text}'
                                                                  ,[Сред.балл] = '{textBox11.Text}'
                                                             WHERE [id] = {dataGridView1.CurrentRow.Cells[0].Value}"); LoadGrid();
                                                                        MessageBox.Show("Запись успешно изменена", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                                                    }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = Class.Database.Query($@"DELETE FROM [dbo].[uspevaemost f]
      WHERE id = {dataGridView1.CurrentRow.Cells[0].Value}"); LoadGrid();
            MessageBox.Show("Запись успешно удалена", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Ведомость успешно сформирована", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Add();
            Excel.Worksheet worksheet = workbook.Sheets[1];
            worksheet.Name = "Отчёт";
            worksheet.Cells[2, 3] = "Ведомость успеваемости учащихся";
            Excel.Range rng1 = worksheet.Range[worksheet.Cells[2, 3], worksheet.Cells[2, 3]];
            rng1.Cells.Font.Name = "Times New Roman";
            rng1.Cells.Font.Size = 24;
            rng1.Font.Bold = true;
            rng1.Cells.Font.Color = ColorTranslator.ToOle(Color.Black);
            worksheet.Cells[4, 1] = "№" +
                                   "п/п";
            worksheet.Columns[1].ColumnWidth = 7;
            worksheet.Cells[4, 2] = "Фамилия,Имя";
            worksheet.Columns[2].ColumnWidth = 30;
            worksheet.Cells[4, 3] = "Класс";
            worksheet.Columns[3].ColumnWidth = 10;
            worksheet.Cells[4, 4] = "Уваж.пропуски";
            worksheet.Columns[4].ColumnWidth = 15;
            worksheet.Cells[4, 5] = "Неуваж.пропуски";
            worksheet.Columns[5].ColumnWidth = 16;
            worksheet.Cells[4, 6] = "Математика";
            worksheet.Columns[6].ColumnWidth = 15;
            worksheet.Cells[4, 7] = "Русский язык";
            worksheet.Cells[7].ColumnWidth = 15;
            worksheet.Cells[4, 8] = "Английский язык";
            worksheet.Columns[8].ColumnWidth = 15;
            worksheet.Cells[4, 9] = "Литература";
            worksheet.Columns[9].ColumnWidth = 15;
            worksheet.Cells[4, 10] = "Физкультура";
            worksheet.Columns[10].ColumnWidth = 15;
            worksheet.Cells[4, 11] = "Физика";
            worksheet.Columns[11].ColumnWidth = 15;
            worksheet.Cells[4, 12] = "Средний балл";
            worksheet.Columns[12].ColumnWidth = 15;
            Excel.Range rng2 = worksheet.Range[worksheet.Cells[4, 1], worksheet.Cells[4, 12]];
            rng2.Font.Bold = true;
            string SqlText = "Select * from [uspevaemost f]";
            SqlDataAdapter adapter = new SqlDataAdapter(SqlText, Class.Database.constr);
            DataTable table = new DataTable();
            adapter.Fill(table);
            int i = 5;
            foreach (DataRow row in table.Rows)
            {
                worksheet.Cells[i, 1] = row["id"];
                worksheet.Cells[i, 2] = row["ФИО"];
                worksheet.Cells[i, 3] = row["Класс"];
                worksheet.Cells[i, 4] = row["Пропуски по уваж"];
                worksheet.Cells[i, 5] = row["Пропуски по неуваж"];
                worksheet.Cells[i, 6] = row["Математика"];
                worksheet.Cells[i, 7] = row["Русский язык"];
                worksheet.Cells[i, 8] = row["Английский язык"];
                worksheet.Cells[i, 9] = row["Литература"];
                worksheet.Cells[i, 10] = row["Физкультура"];
                worksheet.Cells[i, 11] = row["Физика"];
                worksheet.Cells[i, 12] = row["Сред.балл"];
                i++;
                Excel.Range rng3 = worksheet.Range[worksheet.Cells[4, 1], worksheet.Cells[i - 1, 12]];
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

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = Class.Database.Query(@"SELECT [id] as Номер
      ,[ФИО]
      ,[Класс]
      ,[Пропуски по уваж]
      ,[Пропуски по неуваж]
      ,[Математика]
      ,[Русский язык]
      ,[Английский язык]
      ,[Литература]
      ,[Физкультура]
      ,[Физика]
      ,[Сред.балл]
  FROM [dbo].[uspevaemost f] WHERE [ФИО]  LIKE '" + textBox12.Text + "%'");
        }

        private void обновитьТаблицуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            UpdateGrid();
        }
        private void UpdateGrid(bool showMessage = true)
        {
            DataTable updated = new DataTable();

            if (Class.Database.ExecuteSqlCommand(@"SELECT [id] as Номер
      ,[ФИО]
      ,[Класс]
      ,[Пропуски по уваж]
      ,[Пропуски по неуваж]
      ,[Математика]
      ,[Русский язык]
      ,[Английский язык]
      ,[Литература]
      ,[Физкультура]
      ,[Физика]
      ,[Сред.балл]
  FROM [dbo].[uspevaemost f]", updated))
            {
                dataGridView1.DataSource = updated;
                dataGridView1.Columns[0].Visible = false;
                if (showMessage)
                    MessageBox.Show("Данные обновлены.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            updated.Dispose();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Menu f = new Menu();
            f.Show();
            this.Hide();
        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }
    }
}
