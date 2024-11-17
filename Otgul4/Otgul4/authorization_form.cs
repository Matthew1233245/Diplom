using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Otgul4
{
    public partial class authorization_form : Form
    {
        /*OleDbConnection con;
        OleDbCommand cmd;
        OleDbDataReader dr;*/

        private string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=authorization_DB.mdb";




        public authorization_form()
        {
            InitializeComponent();

        }
        private void authorization_form_Load(object sender, EventArgs e)
        {
            

            //con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=authorization_DB.mdb");
            //cmd = new OleDbCommand();
            //dataTable = new DataTable();
            //con.Open();


        }

        private void button1_Click(object sender, EventArgs e)
        {


            //string login = textBox_login.Text;
            //string password = textBox_password.Text;
            //object status;

            //con.Open();
            //cmd.Connection = con;
            ///*cmd.CommandText = $"tab_num LIKE '%textBox1.Text.Trim()%'";*/
            //cmd.CommandText = $"SELECT Status FROM Authorization WHERE Login = '{login}'";
            //dr = cmd.ExecuteReader();
            //dr.Read();

            //status = dr["Status"];

            //if(status.ToString() == "Пользователь")
            //{
            //    MessageBox.Show("Пользователь",
            //            "Информация",
            //            MessageBoxButtons.OK,
            //            MessageBoxIcon.Information);
            //}
            //else if(status.ToString() == "Администратор")
            //{
            //    MessageBox.Show("Администратор",
            //            "Информация",
            //            MessageBoxButtons.OK,
            //            MessageBoxIcon.Information);
            //}

            //con.Close();

            string login = textBox_login.Text;
            string password = textBox_password.Text;

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();
                string query = "SELECT Status FROM Authorization WHERE Login = @Login AND Password = @Password";
                using (OleDbCommand command = new OleDbCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Login", login);
                    command.Parameters.AddWithValue("@Password", password);

                    using (OleDbDataReader reader = command.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            string status = reader["Status"].ToString();
                            if (status == "Пользователь")
                            {
                                MessageBox.Show("Это пользователь");
                            }
                            else if (status == "Администратор")
                            {
                                MessageBox.Show("Это администратор");
                            }
                        }
                        else
                        {
                            MessageBox.Show("Неверный логин или пароль");
                        }
                    }
                }
            }
        }

        
    }
}
