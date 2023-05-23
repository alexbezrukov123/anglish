using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace Успеваемость_учащихся.Forms
{
    public partial class Ucheniki : Form
    {
        public Ucheniki()
        {
            InitializeComponent();
        }

        private void Ucheniki_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "uspevaemotsDataSet.Class1". При необходимости она может быть перемещена или удалена.
            this.class1TableAdapter.Fill(this.uspevaemotsDataSet.Class1);
            LoadGrid();
        }
        private void LoadGrid()
        {
            dataGridView1.DataSource = Class.Database.Query(@"SELECT [id] as 'Номер п/п'
      ,[ФИО] as 'Фамилия, Имя'
      ,[Дата рождения]
      ,[Класс]
  FROM [dbo].[Ucheniki]");
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = Class.Database.Query($@"INSERT INTO [dbo].[Ucheniki]
           ([id]
           ,[ФИО]
           ,[Дата рождения]
           ,[Класс])
     VALUES
           ('{textBox1.Text}',
            '{textBox2.Text}',
            '{dateTimePicker1.Value}',
            '{comboBox1.Text}')"); LoadGrid();
            MessageBox.Show("Запись добавлена", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = Class.Database.Query($@"UPDATE [dbo].[Ucheniki]
   SET 
       [id] =   '{textBox1.Text}'
      ,[ФИО] = '{textBox2.Text}'
      ,[Дата рождения] = '{dateTimePicker1.Value}'
      ,[Класс] = '{comboBox1.Text}'
 WHERE [id] = {dataGridView1.CurrentRow.Cells[0].Value}"); LoadGrid();
            MessageBox.Show("Запись изменена", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = Class.Database.Query($@"DELETE FROM [dbo].[Ucheniki]
      WHERE [id] = {dataGridView1.CurrentRow.Cells[0].Value}"); LoadGrid();
            MessageBox.Show("Запись удалена", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Отчет сформирован", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Add();
            Excel.Worksheet worksheet = workbook.Sheets[1];
            worksheet.Name = "Отчёт";
            worksheet.Cells[2, 2] = "Список учащихся";
            Excel.Range rng1 = worksheet.Range[worksheet.Cells[2, 2], worksheet.Cells[2, 2]];
            rng1.Cells.Font.Name = "Times New Roman";
            rng1.Cells.Font.Size = 24;
            rng1.Font.Bold = true;
            rng1.Cells.Font.Color = ColorTranslator.ToOle(Color.Green);
            worksheet.Cells[4, 1] = "Номер п/п";
            worksheet.Columns[1].ColumnWidth = 10;
            worksheet.Cells[4, 2] = "ФИО";
            worksheet.Columns[2].ColumnWidth = 35;
            worksheet.Cells[4, 3] = "Дата рождения";
            worksheet.Cells[3].ColumnWidth = 18;
            worksheet.Cells[4, 4] = "Класс";
            worksheet.Columns[4].ColumnWidth = 10;
            Excel.Range rng2 = worksheet.Range[worksheet.Cells[4, 1], worksheet.Cells[4, 4]];
            rng2.Font.Bold = true;
            string SqlText = "Select * from Ucheniki";
            SqlDataAdapter adapter = new SqlDataAdapter(SqlText, Class.Database.constr);
            DataTable table = new DataTable();
            adapter.Fill(table);
            int i = 5;
            foreach (DataRow row in table.Rows)
            {
                worksheet.Cells[i, 1] = row["id"];
                worksheet.Cells[i, 2] = row["ФИО"];
                worksheet.Cells[i, 3] = row["Дата рождения"];
                worksheet.Cells[i, 4] = row["Класс"];
                i++;
                Excel.Range rng3 = worksheet.Range[worksheet.Cells[4, 1], worksheet.Cells[i - 1, 4]];
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

        private void Button5_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = Class.Database.Query(@"SELECT [id] as Номер
      ,[ФИО]
      ,[Дата рождения]
      ,[Класс]
  FROM [dbo].[Ucheniki] WHERE [ФИО]  LIKE '" + textBox6.Text + "%' or [Дата рождения] LIKE '" + textBox6.Text + "%' or [Класс] LIKE '" + textBox6.Text + "%' or [id] LIKE '" + textBox6.Text + "%'");
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

            if (Class.Database.ExecuteSqlCommand(@"SELECT [id] as Номер
      ,[ФИО] as 'Фамилия, Имя'
      ,[Дата рождения]
      ,[Класс]
  FROM [dbo].[Ucheniki]", updated))
            {
                dataGridView1.DataSource = updated;
                dataGridView1.Columns[0].Visible = false;
                DataTable cbData = new DataTable();
                if (Class.Database.ExecuteSqlCommand("SELECT id, [Название класса] FROM Class1", cbData))
                {
                    comboBox1.ValueMember = "id";
                    comboBox1.DisplayMember = "Название класса";
                    comboBox1.DataSource = cbData;
                }
                if (showMessage)
                    MessageBox.Show("Данные обновлены.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            updated.Dispose();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox1.DataSource = Class.Database.Query(@"SELECT 
       [id] as Номер
      ,[Название класса]
  FROM [dbo].[Class1]");
            comboBox1.DisplayMember = "Название класса";// столбец для отображения
            comboBox1.ValueMember = "id";//столбец с id
        }
        }
    }

