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
    public partial class Menu : Form
    {
        public Menu()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Ucheniki f = new Ucheniki();
            f.Show();
            this.Hide();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Prepod f = new Prepod();
            f.Show();
            this.Hide();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Uspevaemost f = new Uspevaemost();
            f.Show();
            this.Hide();
        }

        private void Menu_Load(object sender, EventArgs e)
        {
            timer1.Interval = 10;
            timer1.Enabled = true;
            timer1.Start();
        }

        private void timer1_Tick(object sender, EventArgs e)
        { 
            label1.Text = DateTime.Now.ToLongTimeString();
            label2.Text = DateTime.Now.ToLongDateString();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Klass f = new Klass();
            f.Show();
            this.Hide();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Statistika f = new Statistika();
            f.Show();
            this.Hide();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void справкаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Данное приложение предназначено для работы с базой данных.\nПодробное описание вы можете посмотреть во вкладке Руководство пользователя.\n" +
         "Безруков А.К студент 25-тп группа\nПочта для связи bezrukov@mail.com \n 2023 г.", "Справка", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void руководствоПользователяToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Для управления используются кнопки представленные на форме - Добавить запись (добавление записи в таблицу), Удалить запись (Удаление записи из таблицы)" +
               "Изменить запись (Изменение записи таблицы), Выход (Закрытие приложения) элементы TextBox (строки для ввода данных) с помощью которых " +
               "осуществляется взаимодействие кнопок с вашей базой данных.\n\nДля работы с БД вам необходимо заполнить поля для ввода данными об учащемся и нажать кнопку Добавить запись для " +
               "добавления новой записи в таблиц либо двойным кликом выбрать существующую запись для удаления (кнопка Удалить запись) или изменения этой записи (Изменить запись). ", "Руководство пользователя", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
