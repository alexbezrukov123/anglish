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
    public partial class Avtoriz : Form
    {
        public Avtoriz()
        {
            InitializeComponent();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == false)
                textBox2.UseSystemPasswordChar = true;
            
        }

        private void btnEnter_Click(object sender, EventArgs e)
        {
            DataTable tb = new DataTable();

            bool authSuccess = Class.Database.ExecuteSqlCommand($@"SELECT Login,Password
                                                            FROM Avtoriz WHERE (Login = N'{comboBox1.Text}') AND (Password = N'{textBox2.Text}')", tb) && tb.Rows.Count > 0;//проверка введёныых данных

            if (!authSuccess)
            {
                MessageBox.Show("Неверный логин или пароль", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            this.Hide();
            Forms.Menu mn = new Forms.Menu();
            mn.Show(); //закрытие окна
        }

        private void Avtoriz_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Regestr f = new Regestr();
            f.Show();
            this.Hide();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
