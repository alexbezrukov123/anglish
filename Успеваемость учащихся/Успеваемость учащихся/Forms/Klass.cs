using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Успеваемость_учащихся.Forms
{
    public partial class Klass : Form
    {
        public Klass()
        {
            InitializeComponent();
        }

        private void Klass_Load(object sender, EventArgs e)
        {
            LoadGrid();
        }
        private void LoadGrid()
        {
            dataGridView1.DataSource = Class.Database.Query(@"SELECT 
       [id] as Номер
      ,[Название класса]
  FROM [dbo].[Class1]");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = Class.Database.Query($@"INSERT INTO [dbo].[Class1]
           ([id]
           ,[Название класса])
     VALUES

           ('{textBox1.Text}',
            '{textBox2.Text}')"); LoadGrid();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Изменить запись?\n"
                       + "Название группы: " + textBox2.Text.ToString(),
                       "сообщение",
                           MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
            if (result == DialogResult.OK)
            {
                dataGridView1.DataSource = Class.Database.Query($@"UPDATE [dbo].[Class1]
                                 SET [id] = '{textBox1.Text}'
                                 ,[Название класса] = '{textBox2.Text}'
                                 WHERE   id =" + dataGridView1.CurrentRow.Cells[0].Value);

                LoadGrid();
            }
    }

        private void button3_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = Class.Database.Query($@"DELETE FROM [dbo].[Class1]
      WHERE id = {dataGridView1.CurrentRow.Cells[0].Value}"); LoadGrid();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Menu f = new Menu();
            f.Show();
            this.Hide();
        }
    }
}

