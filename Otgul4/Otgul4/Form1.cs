using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace Otgul4
{
    public partial class Form1 : Form
    {
        OleDbConnection con;
        OleDbCommand cmd;
        OleDbDataReader dr;
        //private DataTable dataTable;
        

        string tab_num;

        //private OleDbDataAdapter dataAdapter;
        //private DataTable dataTable;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "dO_Database1DataSet.SpecialDays". При необходимости она может быть перемещена или удалена.
            this.specialDaysTableAdapter.Fill(this.dO_Database1DataSet.SpecialDays);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "dO_Database1DataSet.People". При необходимости она может быть перемещена или удалена.
            this.peopleTableAdapter.Fill(this.dO_Database1DataSet.People);


            con = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DO_Database1.mdb");
            cmd = new OleDbCommand();
            //dataTable = new DataTable();
            con.Open();
            mainTableTableAdapter.Connection = con;
            this.mainTableTableAdapter.Fill(this.dO_Database1DataSet.MainTable);

            dataGridView1.DataSource = this.dO_Database1DataSet.MainTable;
            dataGridView2.DataSource = this.dO_Database1DataSet.MainTable;
            dataGridView3.DataSource = this.dO_Database1DataSet.People;
            dataGridView4.DataSource = this.dO_Database1DataSet.SpecialDays;

            textBox1.KeyDown += new KeyEventHandler(textBox1_KeyDown);
            //Руссификация dataGridView1
            dataGridView1.Columns[1].HeaderText = "Табельный номер";
            dataGridView1.Columns[2].HeaderText = "Подразделение";
            dataGridView1.Columns[3].HeaderText = "Фамилия";
            dataGridView1.Columns[4].HeaderText = "Имя";
            dataGridView1.Columns[5].HeaderText = "Отчество";
            dataGridView1.Columns[6].HeaderText = "Работа/Отгул";
            dataGridView1.Columns[7].HeaderText = "Дата события";
            dataGridView1.Columns[8].HeaderText = "Часы с";
            dataGridView1.Columns[9].HeaderText = "Часы по";
            dataGridView1.Columns[10].HeaderText = "Фактическое время";
            dataGridView1.Columns[11].HeaderText = "Комментарий";
            dataGridView1.Columns[12].HeaderText = "Дата заявления";
            dataGridView1.Columns[13].HeaderText = "Дата принятия заявления";
            dataGridView1.Columns[14].HeaderText = "Статус";
            dataGridView1.Columns["Time_from"].DefaultCellStyle.Format = "t";
            dataGridView1.Columns["Time_end"].DefaultCellStyle.Format = "t";
            dataGridView1.Columns["Actual_time"].DefaultCellStyle.Format = "t";
            //Руссификация dataGridView2
            dataGridView2.Columns[1].HeaderText = "Табельный номер";
            dataGridView2.Columns[2].HeaderText = "Подразделение";
            dataGridView2.Columns[3].HeaderText = "Фамилия";
            dataGridView2.Columns[4].HeaderText = "Имя";
            dataGridView2.Columns[5].HeaderText = "Отчество";
            dataGridView2.Columns[6].HeaderText = "Работа/Отгул";
            dataGridView2.Columns[7].HeaderText = "Дата события";
            dataGridView2.Columns[8].HeaderText = "Часы с";
            dataGridView2.Columns[9].HeaderText = "Часы по";
            dataGridView2.Columns[10].HeaderText = "Фактическое время";
            dataGridView2.Columns[11].HeaderText = "Комментарий";
            dataGridView2.Columns[12].HeaderText = "Дата заявления";
            dataGridView2.Columns[13].HeaderText = "Дата принятия заявления";
            dataGridView2.Columns[14].HeaderText = "Статус";
            dataGridView2.Columns["Actual_time"].DefaultCellStyle.Format = "t";
            //руссификация dataGridView3
            dataGridView3.Columns[1].HeaderText = "Фамилия";
            dataGridView3.Columns[2].HeaderText = "Имя";
            dataGridView3.Columns[3].HeaderText = "Отчество";
            dataGridView3.Columns[4].HeaderText = "Табельный номер";
            dataGridView3.Columns[5].HeaderText = "Подразделение";
            dataGridView3.Columns[6].HeaderText = "Должность";
            dataGridView3.Columns[7].HeaderText = "Время с";
            dataGridView3.Columns[8].HeaderText = "Время по";
            dataGridView3.Columns[9].HeaderText = "Время с (пятница)";
            dataGridView3.Columns[10].HeaderText = "Время по (пятница)";
            dataGridView3.Columns[11].HeaderText = "ФИО (для подписи)";
            dataGridView3.Columns[12].HeaderText = "ФИО (от кого)";
            dataGridView3.Columns[13].HeaderText = "Статус";
            dataGridView3.Columns["M_T_Begin"].DefaultCellStyle.Format = "t";
            dataGridView3.Columns["M_T_End"].DefaultCellStyle.Format = "t";
            dataGridView3.Columns["Fri_Begin"].DefaultCellStyle.Format = "t";
            dataGridView3.Columns["Fri_End"].DefaultCellStyle.Format = "t";
            //руссификация dataGridView4
            dataGridView4.Columns[1].HeaderText = "Дата";
            dataGridView4.Columns[2].HeaderText = "Продолжительность рабочего дня";
            dataGridView4.Columns["SpecialLenght"].DefaultCellStyle.Format = "t";
            dataGridView1_Filter();
            fill_comboBox1();

            con.Close();

            CheckIfCellAreEmpty();

            if (!textBox7.Enabled & !textBox2.Enabled & !textBox5.Enabled & !textBox6.Enabled & 
                !textBox12.Enabled & !textBox3.Enabled & !textBox4.Enabled & !textBox8.Enabled &
                !textBox16.Enabled)
            {
                textBox7.BackColor = Color.White;
                textBox2.BackColor = Color.White;
                textBox5.BackColor = Color.White;
                textBox6.BackColor = Color.White;
                textBox12.BackColor = Color.White;
                textBox3.BackColor = Color.White;
                textBox4.BackColor = Color.White;
                textBox8.BackColor = Color.White;
                textBox16.BackColor = Color.White;

            }

            /*if (dataGridView1.Rows.Count == 0)
            {
                button2.Visible = false;
            }
            else
            {
                button2.Visible = true;
            }*/

            //Всплывающие подсказки виджетов
            ToolTip t = new ToolTip();
            t.SetToolTip(checkBox2, "Нажмите для изменения даты подачи заявления");
            t.SetToolTip(textBox1, "Введите табельный номер");
            t.SetToolTip(button4, "Нажмите для заполнения полей \"Подразделение\" и \"ФИО\"");
            t.SetToolTip(button2, "Аннулирование выделенной строки");
            t.SetToolTip(button3, "Очистка всех полей");
            t.SetToolTip(button1, "Добавление новой записи");

        }
        private void dataGridView1_Filter() //Фильтр dataGridView1
        {
            tab_num = textBox1.Text;
            string query = "SELECT * FROM MainTable WHERE Tab_num = @tab_num";
            using (OleDbCommand cmd = new OleDbCommand(query, con))
            {
                cmd.Parameters.AddWithValue("@tab_num", tab_num);
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(cmd))
                {
                    DataTable dataTable = new DataTable();
                    adapter.Fill(dataTable);
                    dataGridView1.DataSource = dataTable;
                }
            }
        }

        private void fill_comboBox1() //Заполнение поля "Подразделение" на странице добавления сотрудника
        {
            cmd.Connection = con;
            cmd.CommandText = "SELECT * FROM Department";
            dr = cmd.ExecuteReader();
            comboBox1.Items.Clear();

            while (dr.Read())
            {
                comboBox1.Items.Add(dr["Num_Dep"]);
            }
        }

        private bool IsCellEmpty(DataGridViewCell cell)
        {
            return cell.Value == null || cell.Value == DBNull.Value || string.IsNullOrEmpty(cell.Value.ToString());
        }

        private void CheckIfCellAreEmpty()
        {
            bool allCellsEmpty = true;

            foreach(DataGridViewRow row in dataGridView1.Rows)
            {
                foreach(DataGridViewCell cell in row.Cells)
                {
                    if (!IsCellEmpty(cell))
                    {
                        allCellsEmpty = false;
                        break;
                    }
                }
                if (!allCellsEmpty)
                {
                    break;
                }
            }

            if (allCellsEmpty)
            {
                button2.Enabled = false;
                button2.BackColor = Color.LightGray;
            }
            else
            {
                button2.Enabled = true;
                button2.BackColor = Color.LightCoral;
            }
        }

        /*private void dataGridView2_Filter_TabNum()
        {
            tab_num = textBox9.Text;
            string query = "SELECT * FROM MainTable WHERE Tab_num = @tab_num";
            using (OleDbCommand cmd = new OleDbCommand(query, con))
            {
                cmd.Parameters.AddWithValue("@tab_num", tab_num);
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(cmd))
                {
                    DataTable dataTable = new DataTable();
                    adapter.Fill(dataTable);
                    dataGridView2.DataSource = dataTable;
                }
            }
        }*/
        /*private void dataGridView2_Filter_ProjectStatus()
        {
            string query = "SELECT * FROM MainTable WHERE Status = 'Проект'";
            using (OleDbCommand cmd = new OleDbCommand(query, con))
            {
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(cmd))
                {
                    DataTable dataTable = new DataTable();
                    adapter.Fill(dataTable);
                    dataGridView2.DataSource = dataTable;
                }
            }
        }
        private void finaleTime(string tabNum, string WorkNotWork)
        {
            object ActualTime;
            con.Open();
            cmd.Connection = con;
            cmd.CommandText = $"SELECT Actual_time FROM MainTable WHERE Tab_Num = '{tabNum}' AND Work_Not_work = '{WorkNotWork}'";
            dr = cmd.ExecuteReader();
            dr.Read();

            ActualTime = dr["Actual_time"];
            textBox3.Text = ActualTime.ToString();

            con.Close();

            
        }*/

/////////////////////////////////////////////////////////////////////////////////

        private void button4_Click_1(object sender, EventArgs e) //Применить
        {
            try
            {

                object surname, first_name, second_name, department_number;
                


                tab_num = textBox1.Text;
                if (!String.IsNullOrWhiteSpace(tab_num))
                {
                    //textBox1.Text = String.Empty;
                    textBox2.Text = String.Empty;
                    textBox5.Text = String.Empty;
                    textBox6.Text = String.Empty;
                    ///////////////////////////////////////////////////////////////////////////////// Заполнение полей ФИО и подразделение
                    con.Open();
                    cmd.Connection = con;
                    /*cmd.CommandText = $"tab_num LIKE '%textBox1.Text.Trim()%'";*/
                    cmd.CommandText = $"SELECT Surname, First_Name, Second_Name, Department_Number FROM People WHERE Tab_Num = '{tab_num}'";
                    dr = cmd.ExecuteReader();
                    dr.Read();




                    department_number = dr["Department_Number"];
                    surname = dr["Surname"];
                    first_name = dr["First_Name"];
                    second_name = dr["Second_Name"];
                    


                    textBox7.Text = department_number.ToString();
                    textBox2.Text = surname.ToString();
                    textBox5.Text = first_name.ToString();
                    textBox6.Text = second_name.ToString();
                    
                    con.Close();

                    dataGridView1_Filter(); //Фильтр грида

                    CheckIfCellAreEmpty();
                    ///////////////////////////////////////////////////////////////////////////////// 

                    /*con.Open();
                    cmd.Connection = con;
                    cmd.CommandText = $"SELECT Work_Not_work FROM MainTable WHERE Tab_Num = '{tab_num}'";
                    dr = cmd.ExecuteReader();
                    dr.Read();

                    work_not_work = dr["Work_Not_work"];

                    string work_not_work_result = work_not_work.ToString();
                    
                    con.Close();

                    if (work_not_work_result == "Отгул")
                    {
                        int overtimeHours = 0;
                        string query = "SELECT Actual_time FROM MainTable WHERE Tab_Num = @tab_num'";
                        using (cmd = new OleDbCommand(query, con))
                        {
                            cmd.Parameters.AddWithValue("@tab_num", tab_num);
                            con.Open();
                            
                            using(OleDbDataReader reader = cmd.ExecuteReader())
                            {
                                while(reader.Read())
                                {
                                    overtimeHours += reader.GetInt32(0);
                                }
                            }
                        }
                        textBox4.Text = overtimeHours.ToString();
                    }*/

                    /*con.Open();
                    //часы переработки
                    string queryOvertime = "SELECT SUM(CDate(Actual_time)) FROM MainTable WHERE Tab_num = @Tab_num AND Work_Not_work = 'Работа'";
                    using (cmd = new OleDbCommand(queryOvertime, con))
                    {
                        cmd.Parameters.AddWithValue("@Tab_num", tab_num);
                        object overtimeResult = cmd.ExecuteScalar();
                        double totalOvertime = overtimeResult != null && overtimeResult != DBNull.Value ? ((TimeSpan)overtimeResult).TotalHours : 0;
                        textBox3.Text = totalOvertime.ToString();
                    }
                    //часы отгулов
                    string queryLeave = "SELECT SUM(CDate(Actual_time)) FROM MainTable WHERE Tab_num = @Tab_num AND Work_Not_work = 'Отгул'";
                    using (cmd = new OleDbCommand(queryLeave, con))
                    {
                        cmd.Parameters.AddWithValue("@Tab_num", tab_num);
                        object leaveResult = cmd.ExecuteScalar();
                        double totalLeave = leaveResult != null && leaveResult != DBNull.Value ? ((TimeSpan)leaveResult).TotalHours : 0;
                        textBox4.Text = totalLeave.ToString();
                    }
                    //Вычисление дельты
                    double difference = double.Parse(textBox3.Text) - double.Parse(textBox4.Text);
                    textBox8.Text = difference.ToString();*/


                    /*int totalOvertime = 0;
                    int totalLeaveHours = 0;

                    string query = "SELECT * FROM MainTable WHERE Tab_num = @Tab_num";
                    cmd = new OleDbCommand(query, con);
                    cmd.Parameters.AddWithValue("@Tab_num", tab_num);

                    con.Open();
                    dr = cmd.ExecuteReader();
                    while (dr.Read())
                    {
                        int overtimeHours = Convert.ToInt32(dr["Work_Not_work"].ToString().Split(':')[0]);
                        totalOvertime += overtimeHours;

                        if(dr["Work_Not_work"].ToString() == "Отгул")
                        {
                            int leaveHours = Convert.ToInt32(dr["Actual_time"].ToString().Split(':')[0]);
                            totalLeaveHours += leaveHours;
                        }
                    }
                    con.Close();
                    textBox3.Text = totalOvertime.ToString();
                    textBox4.Text = totalLeaveHours.ToString();
                    textBox8.Text = (totalOvertime - totalLeaveHours).ToString();*/

                }
                else if (String.IsNullOrWhiteSpace(tab_num))
                {
                    MessageBox.Show("Для заполнения полей введите табельный номер",
                        "Информация",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }

            }
            catch 
            {
                MessageBox.Show(
                    "Такого табельного номера не существует",
                    "Предупреждение",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
            finally
            {
                con.Close();
            }
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e) //Обработчик нажатия Enter для Таб.№ у пользователя
        {
            if(e.KeyCode == Keys.Enter)
            {
                button4_Click_1(sender, e);
            }
            
        }

        private void button1_Click_1(object sender, EventArgs e) //Добавить
        {
            FillingFile();
            try
            {
                if (!String.IsNullOrWhiteSpace(comboBox3.Text) & !String.IsNullOrWhiteSpace(comboBox4.Text) & !String.IsNullOrWhiteSpace(textBox1.Text))
                {
                    //Заполнение полей БД
                    cmd.Connection = con;
                    cmd.CommandText = "INSERT INTO MainTable(Tab_num, Department, Surname, First_Name, Second_Name, Work_Not_work, Event_date, Time_from, Time_end, Actual_time, Comment, Application_date, Status) Values (@tab_num, @department, @surname, @first_Name, @second_Name, @work_Not_work, @event_date, @time_from, @time_end, @actual_time, @comment, @application_date, @status)";
                    con.Open();
                    cmd.Parameters.AddWithValue("@tab_num", textBox1.Text);
                    cmd.Parameters.AddWithValue("@department", textBox7.Text);
                    cmd.Parameters.AddWithValue("@surname", textBox2.Text);
                    cmd.Parameters.AddWithValue("@first_Name", textBox5.Text);
                    cmd.Parameters.AddWithValue("@second_Name", textBox6.Text);
                    cmd.Parameters.AddWithValue("@work_Not_work", comboBox3.Text);
                    cmd.Parameters.AddWithValue("@event_date", dateTimePicker1.Value.ToShortDateString());
                    cmd.Parameters.AddWithValue("@time_from", dateTimePicker2.Value.ToShortTimeString());
                    cmd.Parameters.AddWithValue("@time_end", dateTimePicker3.Value.ToShortTimeString());
                    cmd.Parameters.AddWithValue("@actual_time", textBox12.Text);
                    cmd.Parameters.AddWithValue("@comment", comboBox4.Text);
                    cmd.Parameters.AddWithValue("@application_date", dateTimePicker4.Value.ToShortDateString());
                    cmd.Parameters.AddWithValue("@status", "Проект");
                    cmd.ExecuteNonQuery();

                    

                    MessageBox.Show("Данные успешно записаны",
                                    "Сообщение",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Information);
                    

                    this.mainTableTableAdapter.Fill(this.dO_Database1DataSet.MainTable);

                    //Очистка полей после нажатия
                    comboBox3.SelectedIndex = -1;
                    comboBox4.SelectedIndex = -1;

                    dateTimePicker1.Value = DateTime.Now;
                    dateTimePicker2.Value = DateTime.Now;
                    dateTimePicker3.Value = DateTime.Now;
                    dateTimePicker4.Value = DateTime.Now;

                    dataGridView1_Filter();
                    CheckIfCellAreEmpty();


                }
                else if (String.IsNullOrWhiteSpace(comboBox3.Text) & String.IsNullOrWhiteSpace(comboBox4.Text) & String.IsNullOrWhiteSpace(textBox1.Text))
                {
                    MessageBox.Show("Заполните поля \"Отгул/Работа\", \"Комментарий\" и \"Табельный номер\"",
                            "Сообщение",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                }
                else if (String.IsNullOrWhiteSpace(comboBox3.Text) & String.IsNullOrWhiteSpace(comboBox4.Text))
                {
                    MessageBox.Show("Заполните поля \"Отгул/Работа\" и \"Комментарий\"",
                            "Сообщение",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                }
                else if (String.IsNullOrWhiteSpace(comboBox3.Text) & String.IsNullOrWhiteSpace(textBox1.Text))
                {
                    MessageBox.Show("Заполните поля \"Отгул/Работа\" и \"Табельный номер\"",
                            "Сообщение",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                }
                else if (String.IsNullOrWhiteSpace(comboBox4.Text) & String.IsNullOrWhiteSpace(textBox1.Text))
                {
                    MessageBox.Show("Заполните поля \"Комментарий\" и \"Табельный номер\"",
                            "Сообщение",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                }
                else if (String.IsNullOrWhiteSpace(textBox1.Text))
                {
                    MessageBox.Show("Заполните поле \"Табельный номер\"",
                            "Сообщение",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                }
                else if (String.IsNullOrWhiteSpace(comboBox3.Text))
                {
                    MessageBox.Show("Заполните поле \"Отгул/Работа\"",
                            "Сообщение",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                }
                else if (String.IsNullOrWhiteSpace(comboBox4.Text))
                {
                    MessageBox.Show("Заполните поле \"Комментарий\"",
                            "Сообщение",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Ошибка записи" + "\n" + "\n" + ex,
                        "Ошибка",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
            }
            finally
            {
                con.Close();
            }
            
        }

        private void FillingFile()
        {
            if (comboBox3.Text == "Отгул")
            {
                var helper = new WordHelper("shablon_zayavlenia_na_otgul.doc");

                object fio_for_whom, department_number, profession, fio_for_signature;
                con.Open();
                cmd.Connection = con;
                cmd.CommandText = $"SELECT FIO_for_whom, Department_Number, Profession, FIO_for_Signature FROM People WHERE Tab_Num = '{textBox1.Text}'";
                dr = cmd.ExecuteReader();
                dr.Read();

                fio_for_whom = dr["FIO_for_whom"];
                department_number = dr["Department_Number"];
                profession = dr["Profession"];
                fio_for_signature = dr["FIO_for_Signature"];

                //var helper = new WordHelper("shablon_zayavlenia_na_otgul.doc");

                var items = new Dictionary<string, string>
                {
                    {"<fio_for_whom>", fio_for_whom.ToString()},
                    {"<department_number>", department_number.ToString()},
                    {"<profession>", profession.ToString()},
                    {"<fio_for_signature>", fio_for_signature.ToString()},
                    {"<event_date>", dateTimePicker1.Value.ToString("dd.MM.yyyy")},
                    {"<application_date>", dateTimePicker4.Value.ToString("dd.MM.yyyy")},
                    {"<time_from>", dateTimePicker2.Value.ToString("HH:mm")},
                    {"<time_end>", dateTimePicker3.Value.ToString("HH:mm")},
                    {"<comment>", comboBox4.Text},
                };

                helper.Process(items);
                con.Close();
            }
            else if (comboBox3.Text == "Работа")
            {
                var helper = new WordHelper("shablon_zayavlenia_na_pererabotku.doc");

                object fio_for_whom, department_number, profession, fio_for_signature;
                con.Open();
                cmd.Connection = con;
                cmd.CommandText = $"SELECT FIO_for_whom, Department_Number, Profession, FIO_for_Signature FROM People WHERE Tab_Num = '{textBox1.Text}'";
                dr = cmd.ExecuteReader();
                dr.Read();

                fio_for_whom = dr["FIO_for_whom"];
                department_number = dr["Department_Number"];
                profession = dr["Profession"];
                fio_for_signature = dr["FIO_for_Signature"];

                var items = new Dictionary<string, string>
                {
                    {"<fio_for_whom>", fio_for_whom.ToString()},
                    {"<department_number>", department_number.ToString()},
                    {"<profession>", profession.ToString()},
                    {"<fio_for_signature>", fio_for_signature.ToString()},
                    {"<event_date>", dateTimePicker1.Value.ToString("dd.MM.yyyy")},
                    {"<application_date>", dateTimePicker4.Value.ToString("dd.MM.yyyy")},
                    {"<time_from>", dateTimePicker2.Value.ToString("HH:mm")},
                    {"<time_end>", dateTimePicker3.Value.ToString("HH:mm")},
                    {"<comment>", comboBox4.Text},
                };

                helper.Process(items);
                con.Close();
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e) //Активация/деактивация кнопки Аннулировать на вкладке пользователя
        {
            if (e.RowIndex >= 0)
            {
                dataGridView1.Rows[e.RowIndex].Selected = true;
            }

            DataGridViewRow selectedRow = dataGridView1.SelectedRows[0];
            string currentStatus = selectedRow.Cells["Status"].Value.ToString();
            if(currentStatus == "Аннулировано" || currentStatus == "Принято")
            {
                button2.Enabled = false;
                button2.BackColor = Color.LightGray;
            }
            else
            {
                button2.Enabled = true;
                button2.BackColor = Color.LightCoral;
            }

        }

        private void dataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e) //Обработчик выбора строки через любую ячейку
        {
            if (e.RowIndex >= 0)
            {
                dataGridView1.Rows[e.RowIndex].Selected = true;
            }
        }


        /*private void dataGridView_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            if(dataGridView1.Rows.Count == 0)
            {
                button2.Visible = false;
            }
            else
            {
                button2.Visible = true;
            }
        }*/

        private void button3_Click_1(object sender, EventArgs e) //Очистить поля
        {
            textBox1.Text = String.Empty;
            textBox7.Clear();
            textBox2.Clear();
            textBox5.Clear();
            textBox6.Clear();
            comboBox3.SelectedIndex = -1;
            comboBox4.SelectedIndex = -1;

            dateTimePicker1.Value = DateTime.Now;
            dateTimePicker2.Value = DateTime.Now;
            dateTimePicker3.Value = DateTime.Now;
            dateTimePicker4.Value = DateTime.Now;

            dataGridView1_Filter();
            CheckIfCellAreEmpty();
        }

        private void comboBox3_SelectedIndexChanged_1(object sender, EventArgs e) //Выбор отгул/работа
        {
            if (comboBox3.SelectedIndex == 0)
            {
                con.Open();
                cmd.Connection = con;
                cmd.CommandText = "SELECT * FROM Pref";
                dr = cmd.ExecuteReader();
                comboBox4.Items.Clear();

                while (dr.Read())
                {
                    comboBox4.Items.Add(dr["Reason_Otgul"]);
                }
                con.Close();
            }
            else if (comboBox3.SelectedIndex == 1)
            {
                con.Open();
                cmd.Connection = con;
                cmd.CommandText = "SELECT * FROM Pref";
                dr = cmd.ExecuteReader();
                comboBox4.Items.Clear();

                while (dr.Read())
                {
                    comboBox4.Items.Add(dr["Why_Work"]);
                }
                con.Close();
            }

            //con.Open();
            //cmd.Connection = con;
            //cmd.CommandText = "SELECT * FROM Pref";
            //dr = cmd.ExecuteReader();
            //comboBox4.Items.Clear();

            //if (comboBox3.SelectedIndex == 0)
            //{
                

            //    while (dr.Read())
            //    {
            //        comboBox4.Items.Add(dr["Reason_Otgul"]);
            //    }
            //    con.Close();
            //}
            //else if (comboBox3.SelectedIndex == 1)
            //{

                
            //    while (dr.Read())
            //    {
            //        comboBox4.Items.Add(dr["Why_Work"]);
            //    }
            //    con.Close();
            //}
            //con.Close();
        }

        private void button2_Click_1(object sender, EventArgs e) //Удалить
        {
            /*if (dataGridView1.SelectedRows.Count > 0)
            {
                DialogResult result = MessageBox.Show(
                    "Данные в выделенной строке будут аннулированы. Вы уверены?",
                    "Предупреждение",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    DataGridViewRow selectedRow = dataGridView1.SelectedRows[0];
                    int primaryKey = Convert.ToInt32(selectedRow.Cells["№"].Value);

                    //Удаление из dataGridView
                    dataGridView1.Rows.Remove(selectedRow);
                    //Удаление из БД
                    string deleteQuery = "DELETE FROM MainTable WHERE № = @PrimaryKey";
                    using (OleDbCommand cmd = new OleDbCommand(deleteQuery, con))
                    {
                        cmd.Parameters.AddWithValue("@PrimaryKey", primaryKey);
                        con.Open();
                        cmd.ExecuteNonQuery();
                        con.Close();
                    }
                    if (dataGridView1.SelectedRows.Count > 0)
                    {
                        DataRowView selectedRow = (DataRowView)dataGridView1.SelectedRows[0].DataBoundItem;
                        if (selectedRow["Status"] != "Принято")
                        {
                            selectedRow["Status"] = "Аннулировано";
                            UpdateDatabase_for_delete(selectedRow);
                        }
                        else
                        {
                            MessageBox.Show("Это заявление уже принято. Обратитесь к Администратору.",
                                            "Информация",
                                            MessageBoxButtons.OK,
                                            MessageBoxIcon.Information);
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Выделите строку для удаления в таблице",
                        "Информация",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
            }*/

            
            if(dataGridView1.SelectedRows.Count > 0)
            {
                DialogResult result = MessageBox.Show(
                    "Данные в выделенной строке будут аннулированы. Вы уверены?",
                    "Предупреждение",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    DataGridViewRow selectedRow = dataGridView1.SelectedRows[0];
                    string currentStatus = selectedRow.Cells["Status"].Value.ToString();

                    switch (currentStatus)
                    {
                        case "Проект":
                            selectedRow.Cells["Status"].Value = "Аннулировано";
                            UpdateDB_StatusDelet(selectedRow.Cells["№"].Value.ToString(), "Аннулировано");
                            break;
                        case "Принято":
                            MessageBox.Show("Это заявление уже принято. Обратитесь к Администратору.",
                                            "Информация",
                                            MessageBoxButtons.OK,
                                            MessageBoxIcon.Information);
                            break;
                        case "Аннулировано":
                            MessageBox.Show("Это заявление уже аннулировано.",
                                            "Информация",
                                            MessageBoxButtons.OK,
                                            MessageBoxIcon.Information);
                            break;
                        default:
                            MessageBox.Show("Неизвестный статус",
                                            "Ошибка",
                                            MessageBoxButtons.OK,
                                            MessageBoxIcon.Error);
                            break;
                    }
                    
                }
            }
            else
            {
                MessageBox.Show("Выберите строку для аннулирования",
                                "Информация",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
            }

        }
        private void UpdateDB_StatusDelet(string id, string newStatus) //Обновление БД при аннулировании
        {
            string query = $"UPDATE MainTable SET Status = @Status WHERE № = @ID";
            using(cmd = new OleDbCommand(query, con))
            {
                cmd.Parameters.AddWithValue("@Status", newStatus);
                cmd.Parameters.AddWithValue("@ID", id);
                con.Open();
                cmd.ExecuteNonQuery();
                con.Close();
            }
        }


        private void dateTimePicker2_ValueChanged_1(object sender, EventArgs e) //Время С
        {
            TimeSpan result = this.dateTimePicker3.Value - this.dateTimePicker2.Value;
            this.textBox12.Text = result.ToString();

            string[] tempArry = textBox12.Text.Split('.');
            textBox12.Text = tempArry[0];
        }

        private void dateTimePicker3_ValueChanged_1(object sender, EventArgs e) //Время По
        {
            dateTimePicker2_ValueChanged_1(sender, e);
        }

        private void button5_Click(object sender, EventArgs e) //Кнопка Принять заявление на странице админа
        {
            /*string query = "UPDATE MainTable SET Status = @status WHERE № = @ID";

            using (cmd = new OleDbCommand(query, con))
            {
                cmd.Parameters.AddWithValue("@status", "Аннулировано");
                cmd.Parameters.AddWithValue("@ID", row["№"]);
                con.Open();
                cmd.ExecuteNonQuery();
                con.Close();
            }*/

            /*DataRowView Row = (DataRowView)dataGridView2.SelectedRows[0].DataBoundItem;
            //Row["Status"] = "Аннулировано";
            DataRowView row = Row;
            string query = "SELECT Status FROM MainTable WHERE № = @ID";
            cmd.Parameters.AddWithValue("@ID", row["№"]);
            if (query != "Аннулировано")
            { }*/

            DataGridViewRow Row = dataGridView2.SelectedRows[0];
            string currentStatus = Row.Cells["Status"].Value.ToString();

            switch (currentStatus)
            {
                case "Проект":
                    if (dataGridView2.SelectedRows.Count > 0)
                    {
                        DataRowView selectedRow = (DataRowView)dataGridView2.SelectedRows[0].DataBoundItem;
                        selectedRow["Acceptance_date"] = dateTimePicker5.Value;

                        DataRowView update_status = (DataRowView)dataGridView2.SelectedRows[0].DataBoundItem;
                        update_status["Status"] = "Принято";

                        UpdateDatabase_AcceptanceDate_and_Status(selectedRow, update_status);
                    }
                    break;
                case "Принято":
                    MessageBox.Show("Это заявление уже принято.",
                                    "Информация",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Information);
                    break;

                case "Аннулировано":
                    MessageBox.Show("Это заявление аннулировано.",
                                    "Информация",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Information);
                    break;

            }


            /*if (dataGridView2.SelectedRows.Count > 0)
                {
                    DataRowView selectedRow = (DataRowView)dataGridView2.SelectedRows[0].DataBoundItem;
                    selectedRow["Acceptance_date"] = dateTimePicker5.Value;

                    DataRowView update_status = (DataRowView)dataGridView2.SelectedRows[0].DataBoundItem;
                    update_status["Status"] = "Принято";

                    UpdateDatabase_AcceptanceDate_and_Status(selectedRow, update_status);
                }*/
            
        }
        private void UpdateDatabase_AcceptanceDate_and_Status(DataRowView row, DataRowView upd_status) //Обновление БД под button5
        {
            string query = "UPDATE MainTable SET Acceptance_date = @acp_date WHERE № = @ID";
            using (cmd = new OleDbCommand(query, con))
            {
                cmd.Parameters.AddWithValue("@acp_date", row["Acceptance_date"]);
                cmd.Parameters.AddWithValue("@ID", row["№"]);
                con.Open();
                cmd.ExecuteNonQuery();
                con.Close();
            }

            string query2 = "UPDATE MainTable SET Status = @up_status WHERE № = @ID";
            using (cmd = new OleDbCommand(query2, con))
            {
                cmd.Parameters.AddWithValue("@up_status", upd_status["Status"]);
                cmd.Parameters.AddWithValue("@ID", upd_status["№"]);
                con.Open();
                cmd.ExecuteNonQuery();
                con.Close();
            }
        }

        private void Admin_DataGridView_Filter_With_TabNum(string query_) //Метод для фильтрации dataGridView на форме админа с введенным таб. №
        {
            using (OleDbCommand cmd = new OleDbCommand(query_, con))
            {
                cmd.Parameters.AddWithValue("@tab_num", tab_num);
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(cmd))
                {
                    DataTable dataTable = new DataTable();
                    adapter.Fill(dataTable);
                    dataGridView2.DataSource = dataTable;
                }
            }
        }

        private void Admin_DataGridView_Filter_Without_TabNum(string query_) //Метод для фильтрации dataGridView на форме админа без введенного таб. №
        {
            using (OleDbCommand cmd = new OleDbCommand(query_, con))
            {
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(cmd))
                {
                    DataTable dataTable = new DataTable();
                    adapter.Fill(dataTable);
                    dataGridView2.DataSource = dataTable;
                }
            }
        }

        private void textBox9_TextChanged(object sender, EventArgs e) //Обработчик вводимого Таб.№ в фильтре у Админа
        {
            dataGridView2Filter();
        }
        private void dataGridView2Filter()
        {
            tab_num = textBox9.Text;

            if (!String.IsNullOrWhiteSpace(tab_num)) //Условия фильтрации при введенном таб.№
            {
                if (!checkBox3.Checked && !checkBox4.Checked && !checkBox5.Checked)
                {
                    string query = "SELECT * FROM MainTable WHERE Tab_num = @tab_num";
                    /*using (OleDbCommand cmd = new OleDbCommand(query, con))
                    {
                        cmd.Parameters.AddWithValue("@tab_num", tab_num);
                        using (OleDbDataAdapter adapter = new OleDbDataAdapter(cmd))
                        {
                            DataTable dataTable = new DataTable();
                            adapter.Fill(dataTable);
                            dataGridView2.DataSource = dataTable;
                        }
                    }*/
                    Admin_DataGridView_Filter_With_TabNum(query);
                }
                else if (checkBox3.Checked && !checkBox4.Checked && !checkBox5.Checked)
                {
                    string query = "SELECT * FROM MainTable WHERE Tab_num = @tab_num AND Status = 'Проект'";
                    Admin_DataGridView_Filter_With_TabNum(query);
                }
                else if (!checkBox3.Checked && checkBox4.Checked && !checkBox5.Checked)
                {
                    string query = "SELECT * FROM MainTable WHERE Tab_num = @tab_num AND Status = 'Принято'";
                    Admin_DataGridView_Filter_With_TabNum(query);
                }
                else if (!checkBox3.Checked && !checkBox4.Checked && checkBox5.Checked)
                {
                    string query = "SELECT * FROM MainTable WHERE Tab_num = @tab_num AND Status = 'Аннулировано'";
                    Admin_DataGridView_Filter_With_TabNum(query);
                }
                else if (checkBox3.Checked && checkBox4.Checked && !checkBox5.Checked)
                {
                    string query = "SELECT * FROM MainTable WHERE Tab_num = @tab_num AND (Status = 'Проект' OR Status = 'Принято')";
                    Admin_DataGridView_Filter_With_TabNum(query);
                }
                else if (!checkBox3.Checked && checkBox4.Checked && checkBox5.Checked)
                {
                    string query = "SELECT * FROM MainTable WHERE Tab_num = @tab_num AND (Status = 'Принято' OR Status = 'Аннулировано')";
                    Admin_DataGridView_Filter_With_TabNum(query);
                }
                else if (checkBox3.Checked && !checkBox4.Checked && checkBox5.Checked)
                {
                    string query = "SELECT * FROM MainTable WHERE Tab_num = @tab_num AND (Status = 'Проект' OR Status = 'Аннулировано')";
                    Admin_DataGridView_Filter_With_TabNum(query);
                }
                else if (checkBox3.Checked && checkBox4.Checked && checkBox5.Checked)
                {
                    string query = "SELECT * FROM MainTable WHERE Tab_num = @tab_num AND (Status = 'Проект' OR Status = 'Принято' OR Status = 'Аннулировано')";
                    Admin_DataGridView_Filter_With_TabNum(query);
                }
            }


            if (String.IsNullOrWhiteSpace(tab_num)) //Фильтрация без введенного таб.№
            {
                if (checkBox3.Checked && !checkBox4.Checked && !checkBox5.Checked)
                {
                    string query = "SELECT * FROM MainTable WHERE Status = 'Проект'";
                    /*using (OleDbCommand cmd = new OleDbCommand(query, con))
                    {
                        using (OleDbDataAdapter adapter = new OleDbDataAdapter(cmd))
                        {
                            DataTable dataTable = new DataTable();
                            adapter.Fill(dataTable);
                            dataGridView2.DataSource = dataTable;
                        }
                    }*/
                    Admin_DataGridView_Filter_Without_TabNum(query);
                }
                else if (!checkBox3.Checked && checkBox4.Checked && !checkBox5.Checked)
                {
                    string query = "SELECT * FROM MainTable WHERE Status = 'Принято'";
                    Admin_DataGridView_Filter_Without_TabNum(query);
                }
                else if (!checkBox3.Checked && !checkBox4.Checked && checkBox5.Checked)
                {
                    string query = "SELECT * FROM MainTable WHERE Status = 'Аннулировано'";
                    Admin_DataGridView_Filter_Without_TabNum(query);
                }
                else if (checkBox3.Checked && checkBox4.Checked && !checkBox5.Checked)
                {
                    string query = "SELECT * FROM MainTable WHERE Status = 'Проект' OR Status = 'Принято'";
                    Admin_DataGridView_Filter_Without_TabNum(query);
                }
                else if (!checkBox3.Checked && checkBox4.Checked && checkBox5.Checked)
                {
                    string query = "SELECT * FROM MainTable WHERE Status = 'Принято' OR Status = 'Аннулировано'";
                    Admin_DataGridView_Filter_Without_TabNum(query);
                }
                else if (checkBox3.Checked && !checkBox4.Checked && checkBox5.Checked)
                {
                    string query = "SELECT * FROM MainTable WHERE Status = 'Проект' OR Status = 'Аннулировано'";
                    Admin_DataGridView_Filter_Without_TabNum(query);
                }
                else if ((checkBox3.Checked && checkBox4.Checked && checkBox5.Checked) || (!checkBox3.Checked && !checkBox4.Checked && !checkBox5.Checked))
                {
                    string query = "SELECT * FROM MainTable WHERE Status = 'Проект' OR Status = 'Принято' OR Status = 'Аннулировано'";
                    Admin_DataGridView_Filter_Without_TabNum(query);
                }
                
            }

        }
        private void button7_Click(object sender, EventArgs e) // Кнопка снять фильтр
        {
            textBox9.Text = String.Empty;
            checkBox3.Checked = false;
            checkBox4.Checked = false;
            checkBox5.Checked = false;
            string query = "SELECT * FROM MainTable WHERE Status = 'Проект' OR Status = 'Принято' OR Status = 'Аннулировано'";
            Admin_DataGridView_Filter_Without_TabNum(query);
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e) //разблокировка времени у админа
        {
            dateTimePicker5.Enabled = checkBox1.Checked;
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e) //разблокировка времени у пользователя
        {
            dateTimePicker4.Enabled = checkBox2.Checked;
        }

        private void UpdateDatabase_for_delete(DataRowView row, string query_, string status_str, string code_) //Обновление dataGridView после аннулирования заявления
        {
            /*string query = "UPDATE MainTable SET Status = @status WHERE № = @ID";

            using (cmd = new OleDbCommand(query, con))
            {
                cmd.Parameters.AddWithValue("@status", "Аннулировано");
                cmd.Parameters.AddWithValue("@ID", row["№"]);
                con.Open();
                cmd.ExecuteNonQuery();
                con.Close();
            }*/

            

            using (cmd = new OleDbCommand(query_, con))
            {
                cmd.Parameters.AddWithValue("@status", status_str);
                cmd.Parameters.AddWithValue("@ID", row[code_]);
                con.Open();
                cmd.ExecuteNonQuery();
                con.Close();
            }
        }

        private void button8_Click(object sender, EventArgs e) //Кнопка Аннулирования на форме админа
        {

            if (dataGridView2.SelectedRows.Count > 0)
            {
                string query = "UPDATE MainTable SET Status = @status WHERE № = @ID";
                string status = "Аннулировано";
                string code = "№";
                DataRowView selectedRow = (DataRowView)dataGridView2.SelectedRows[0].DataBoundItem;
                selectedRow["Status"] = "Аннулировано";
                UpdateDatabase_for_delete(selectedRow, query, status, code);
            }
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e) //Обработчик активации и деактивации кнопок при выборе строк в dataGridView2
        {
            if (e.RowIndex >= 0)
            {
                dataGridView2.Rows[e.RowIndex].Selected = true;
            }

            DataGridViewRow selectedRow = dataGridView2.SelectedRows[0];
            string currentStatus = selectedRow.Cells["Status"].Value.ToString();
            if (currentStatus == "Аннулировано")
            {
                button8.Enabled = false;
                button8.BackColor = Color.LightGray;
            }
            else
            {
                button8.Enabled = true;
                button8.BackColor = Color.LightCoral;
            }

            if (currentStatus == "Принято" || currentStatus == "Аннулировано")
            {
                button5.Enabled = false;
                button5.BackColor = Color.LightGray;
            }
            else
            {
                button5.Enabled = true;
                button5.BackColor = Color.PaleGreen;
            }
        }

        private void dataGridView2_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e) //Обработчик выбора строки при клике на любую ячейку в dataGridView2
        {
            if (e.RowIndex >= 0)
            {
                dataGridView2.Rows[e.RowIndex].Selected = true;
            }
        }

        /////////////////////////////////////////////////////////////////////////////////////////////////// Вкладка Настроек

        private void button9_Click(object sender, EventArgs e) //Кнопка Добавить на вкладке добавления сотрудников
        {
            try
            {
                if (!String.IsNullOrWhiteSpace(comboBox1.Text) & !String.IsNullOrWhiteSpace(textBox10.Text) & !String.IsNullOrWhiteSpace(textBox11.Text) 
                    & !String.IsNullOrWhiteSpace(textBox13.Text) & !String.IsNullOrWhiteSpace(textBox14.Text) & !String.IsNullOrWhiteSpace(textBox15.Text) 
                    & !String.IsNullOrWhiteSpace(textBox16.Text) & !String.IsNullOrWhiteSpace(textBox17.Text))
                {
                    cmd.Connection = con;
                    cmd.CommandText = "INSERT INTO People(Surname, First_Name, Second_Name, Tab_Num, Department_Number, Profession, M_T_Begin, M_T_End, Fri_Begin, Fri_End, FIO_for_Signature, FIO_for_whom, Status) Values (@surname, @first_Name, @second_Name, @tab_Num, @department_Number, @profession, @m_T_Begin, @m_T_End, @fri_Begin, @fri_End, @fIO_for_Signature, @fIO_for_whom, @status)";
                    con.Open();
                    cmd.Parameters.AddWithValue("@surname", textBox10.Text);
                    cmd.Parameters.AddWithValue("@first_Name", textBox11.Text);
                    cmd.Parameters.AddWithValue("@second_Name", textBox13.Text);
                    cmd.Parameters.AddWithValue("@tab_Num", textBox14.Text);
                    cmd.Parameters.AddWithValue("@department_Number", comboBox1.Text);
                    cmd.Parameters.AddWithValue("@profession", textBox15.Text);
                    cmd.Parameters.AddWithValue("@m_T_Begin", dateTimePicker6.Value);
                    cmd.Parameters.AddWithValue("@m_T_End", dateTimePicker7.Value);
                    cmd.Parameters.AddWithValue("@fri_Begin", dateTimePicker8.Value);
                    cmd.Parameters.AddWithValue("@fri_End", dateTimePicker9.Value);
                    cmd.Parameters.AddWithValue("@fIO_for_Signature", textBox16.Text);
                    cmd.Parameters.AddWithValue("@fIO_for_whom", textBox17.Text);
                    cmd.Parameters.AddWithValue("@status", "Работает");
                    cmd.ExecuteNonQuery();

                    //@surname, @first_Name, @second_Name, @tab_Num, @department_Number, @profession, @m-T_Begin, @m-T_End, @fri_Begin, @fri_End, @fIO_for_Signature, @fIO_for_whom, @status

                    MessageBox.Show("Данные успешно записаны",
                                    "Сообщение",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Information);


                    this.peopleTableAdapter.Fill(this.dO_Database1DataSet.People);

                    //Очистка полей после нажатия
                    /*comboBox1.SelectedIndex = -1;

                    textBox10.Text = String.Empty;
                    textBox11.Text = String.Empty;
                    textBox13.Text = String.Empty;
                    textBox14.Text = String.Empty;
                    textBox15.Text = String.Empty;
                    textBox16.Text = String.Empty;
                    textBox17.Text = String.Empty;
                    textBox18.Text = String.Empty;

                    dateTimePicker6.Value = DateTime.Now;
                    dateTimePicker7.Value = DateTime.Now;
                    dateTimePicker8.Value = DateTime.Now;
                    dateTimePicker9.Value = DateTime.Now;*/
                }
                else
                {
                    MessageBox.Show("Все поля должны быть заполнены",
                                    "Предупреждение",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Ошибка записи" + "\n" + "\n" + ex,
                                "Ошибка",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
            }
            finally
            {
                con.Close();
            }

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }
        
        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button10_Click(object sender, EventArgs e) //Кнопка Добавить на вкладке добавления особых дней
        {
            try
            {
                cmd.Connection = con;
                cmd.CommandText = "INSERT INTO SpecialDays(SpecialDay, SpecialLenght) Values (@specialDay, @specialLenght)";
                con.Open();
                cmd.Parameters.AddWithValue("@specialDay", dateTimePicker10.Value.ToShortDateString());
                cmd.Parameters.AddWithValue("@specialLenght", dateTimePicker11.Value.ToShortTimeString());

                cmd.ExecuteNonQuery();



                MessageBox.Show("Данные успешно записаны",
                                "Сообщение",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information);


                this.specialDaysTableAdapter.Fill(this.dO_Database1DataSet.SpecialDays);
            }
            catch (Exception ex)
            {

                MessageBox.Show("Ошибка записи" + "\n" + "\n" + ex,
                                "Ошибка",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
            }
            finally
            {
                con.Close();
            }
        }

        private void button11_Click(object sender, EventArgs e) //Кнопка Удалить на вкладке добавления особых дней
        {
            if (dataGridView4.SelectedRows.Count > 0)
            {
                DialogResult result = MessageBox.Show(
                    "Данные в выделенной строке будут удалены. Вы уверены?",
                    "Предупреждение",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    DataGridViewRow selectedRow = dataGridView4.SelectedRows[0];
                    int primaryKey = Convert.ToInt32(selectedRow.Cells["Код"].Value);

                    //Удаление из dataGridView
                    dataGridView4.Rows.Remove(selectedRow);
                    //Удаление из БД
                    string deleteQuery = "DELETE FROM SpecialDays WHERE Код = @PrimaryKey";
                    using (OleDbCommand cmd = new OleDbCommand(deleteQuery, con))
                    {
                        cmd.Parameters.AddWithValue("@PrimaryKey", primaryKey);
                        con.Open();
                        cmd.ExecuteNonQuery();
                        con.Close();
                    }
                }
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e) //Обработчик клика на checkBox3 в фильтре у админа
        {
            dataGridView2Filter();
        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e) //Обработчик клика на checkBox4 в фильтре у админа
        {
            dataGridView2Filter();
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e) //Обработчик клика на checkBox5 в фильтре у админа
        {
            dataGridView2Filter();
        }

        private void button12_Click(object sender, EventArgs e) //Кнопка "Отправить в архив" на вкладке добавления сотрудников
        {
            if (dataGridView3.SelectedRows.Count > 0)
            {
                string query = "UPDATE People SET Status = @status WHERE Код = @ID";
                string status = "Не работает";
                string code = "Код";
                DataRowView selectedRow = (DataRowView)dataGridView3.SelectedRows[0].DataBoundItem;
                selectedRow["Status"] = "Не работает";
                UpdateDatabase_for_delete(selectedRow, query, status, code);
            }

        }

        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e) //Обработчик активации и деактивации кнопок при выборе строки в dataGridView3
        {
            if(e.RowIndex >= 0)
            {
                dataGridView3.Rows[e.RowIndex].Selected = true;
            }

            DataGridViewRow selectedRow = dataGridView3.SelectedRows[0];
            string currentStatus = selectedRow.Cells["Status"].Value.ToString();
            if (currentStatus == "Не работает")
            {
                button12.Enabled = false;
                button12.BackColor = Color.LightGray;
            }
            else
            {
                button12.Enabled = true;
                button12.BackColor = Color.LightCoral;
            }
        }

        private void dataGridView3_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e) //Обработчик выбора строки в dataGridView3 при клике на любую ячейку
        {
            if (e.RowIndex >= 0)
            {
                dataGridView3.Rows[e.RowIndex].Selected = true;
            }
        }

        private void dataGridView4_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                dataGridView4.Rows[e.RowIndex].Selected = true;
            }
        }

        private void dataGridView4_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                dataGridView4.Rows[e.RowIndex].Selected = true;
            }
        }
        private void FIO_for_signature() //Создание макета фамилии и инициалов для подписи
        {
            string text_surname = textBox10.Text;
            string firstLetter_firstName = textBox11.Text.Length > 0 ? textBox11.Text[0].ToString() : "";
            string firstLetter_secondName = textBox13.Text.Length > 0 ? textBox13.Text[0].ToString() : "";

            textBox16.Text = text_surname + " " + firstLetter_firstName + "." + " " + firstLetter_secondName + ".";
        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {
            FIO_for_signature();
        }

        private void textBox11_TextChanged(object sender, EventArgs e)
        {
            FIO_for_signature();
        }

        private void textBox13_TextChanged(object sender, EventArgs e)
        {
            FIO_for_signature();
        }

        private void textBox18_TextChanged(object sender, EventArgs e)
        {
            
        }


        //private void button13_Click(object sender, EventArgs e)
        //{
        //    /*string query = "SELECT SpecialLenght AS TotalSpecialLenght FROM SpecialDays";

        //    con.Open();
        //    using (cmd = new OleDbCommand(query, con))
        //    {
        //        object result = cmd.ExecuteScalar();
        //        if (result != DBNull.Value && result != null)
        //        {
        //            TimeSpan totalTime = TimeSpan.FromHours(Convert.ToDouble(result));
        //            textBox18.Text = totalTime.ToString(@"hh\:mm");

        //            //double totalDays = Convert.ToDouble(result);
        //            //TimeSpan totalTime = TimeSpan.FromHours(totalDays);
        //            //textBox18.Text = totalTime.ToString(@"hh\:mm");
        //        }
        //        else
        //        {
        //            textBox18.Text = "No data";
        //        }
        //        con.Close();
        //    }*/

        //    string query = "SELECT SUM(SpecialLenght) FROM SpecialDays";
        //    try
        //    {
        //        con.Open();
        //        cmd = new OleDbCommand(query, con);

        //        object result = cmd.ExecuteScalar();
        //        if (result != DBNull.Value && result != null)
        //        {
        //            textBox18.Text = result.ToString();
        //        }
        //        else
        //        {
        //            textBox18.Text = "No data";
        //        }
        //        con.Close();



        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Ошибка: " + ex.Message);
        //    }

        //    /*con.Open();
        //    using (cmd = new OleDbCommand(query, con))
        //    {
        //        object result = cmd.ExecuteScalar();
        //        if (result != DBNull.Value && result != null)
        //        {
        //            DateTime dateTimeValue = Convert.ToDateTime(result);
        //            dateTimePicker12.Value = dateTimeValue;

        //            //double totalDays = Convert.ToDouble(result);
        //            //TimeSpan totalTime = TimeSpan.FromHours(totalDays);
        //            //textBox18.Text = totalTime.ToString(@"hh\:mm");
        //        }
        //        else
        //        {
        //            textBox18.Text = "No data";
        //        }
        //        con.Close();
        //    }*/

        //    /*string query = "SELECT SpecialLenght FROM SpecialDays WHERE SpecialLenght Is Not Null";

        //    con.Open();

        //    using (cmd = new OleDbCommand(query, con))
        //    {
        //        using (dr = cmd.ExecuteReader())
        //        {
        //            TimeSpan totalHours = new TimeSpan();
        //            while (dr.Read())
        //            {
        //                DateTime currentHours = dr.GetDateTime(0);
        //                totalHours = totalHours.Add(currentHours.TimeOfDay);
        //            }
        //            dateTimePicker12.Value = DateTime.Today.Add(totalHours);
        //        }
        //    }

        //    con.Close();*/
        //}


    }
}
