using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Успеваемость_учащихся.Class
{
    class Database
    {
        public static string constr = @"Data Source=62.78.81.19;Initial Catalog=uspevaemots;User ID=25-тпБезруковАК;Password=890694";
        public static DataTable Query(string sql)
        {
            DataTable dt = new DataTable();
            try
            {
                SqlDataAdapter da = new SqlDataAdapter(sql, constr);
                da.Fill(dt);

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка обращения к БД!\nПроверьте вводимые данные \n {ex.Message}", "Уведомление", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);

            }
            return dt;
        }

        static public bool ExecuteSqlCommand(string sql, DataTable toFill)
        {
            //try
            //{
            SqlDataAdapter adapter = new SqlDataAdapter(sql, constr);
            adapter.Fill(toFill);

            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("Ошибка обращения к БД!\nПроверьте вводимые данные.", "Уведомление", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    return false;
            //}

            return true;
        }
    }
}
