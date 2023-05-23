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

namespace Успеваемость_учащихся.Forms
{
    public partial class Regestr : Form
    {
        public Regestr()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Avtoriz f = new Avtoriz();
            f.Show();
            this.Hide();
        }

        private void Regestr_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            var NameUser = textBox1.Text;
            var emailUser = textBox2.Text;
            var passwordUser = textBox3.Text;


            string query = $"insert into Registr(name,email,password) values ('{NameUser}','{emailUser}','{passwordUser}')";

            using (SqlConnection connect = new SqlConnection(Class.Database.constr))
            {
                connect.Open();
                SqlCommand command = new SqlCommand(query, connect);
                if (command.ExecuteNonQuery() == 1)
                {
                    MessageBox.Show("Регистрация успешна!", "Регистрация выполнена", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Forms.Avtoriz auth = new Forms.Avtoriz();
                    auth.Show(); this.Hide();
                }
                else
                {
                    MessageBox.Show("Вы неправильно ввели значения при регистрации!", "Регистрация не выполнена", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }
}
