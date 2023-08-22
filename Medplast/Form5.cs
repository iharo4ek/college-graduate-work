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
namespace Medplast {
    public partial class Form5 : Form {
        private const string STRING_PATTERN = @"[^а-яА-яa-zA-Z]";
        private const string INT_PATTERN = @"[^0-9]";
        DataBase dataBase = DataBase.getInstance();
        User user = User.getInstance();
        Modes mode;
        private string filter = $"";
        private static Form instance;
        private SqlConnection sql;
        public static Form getInstance() {
            if (instance == null) {
                instance = new Form5();
            }
            return instance;
        }
        private void setMode() {
            User user = User.getInstance();
            switch (user.getJobTitle()) {
                case "директор": {
                        this.mode = Modes.READ;
                        break;
                    }
                case "зам директора": {
                        this.mode = Modes.READ;
                        break;
                    }
                case "главный бухгалтер": {
                        this.mode = Modes.READ;
                        break;
                    }
                case "бухгалтер": {
                        this.mode = Modes.READ;
                        break;
                    }
                default: {
                        this.mode = Modes.READWRITE;
                        break;
                    }
            }
        }
        private void getAccess() {
            if (mode == Modes.READ) {
                button1.Visible = false;
                button2.Visible = false;
                button3.Visible = false;
                textBox1.Visible = false;
                textBox3.Visible = false;
                textBox4.Visible = false;
                textBox5.Visible = false;
                textBox6.Visible = false;
                dateTimePicker1.Visible = false;
                comboBox1.Visible = false;
                comboBox2.Visible = false;
                label1.Visible = false;
                label2.Visible = false;
                label3.Visible = false;
                label4.Visible = false;
                label5.Visible = false;
                label6.Visible = false;
                label7.Visible = false;
                label8.Visible = false;
                maskedTextBox1.Visible = false;
                label10.Visible = false;
            }
        }
        private Form5() {
            InitializeComponent();
        }
        private void employees() {
            string query = "";
            if (user.getJobTitle() != "администратор") {
                query = "select id_employee, departmentName as [отдел],jobTitle as [должность], employeeSurname as [фамилия], employeeName as[имя],employeePatronymic as[отчество], bornDate as [дата рождения], passport as [пасспорт] from employees inner join departments on employees.id_department = departments.id_depatment inner join jobTitles on employees.id_jobTitle = jobTitles.id_jobTitle;";
            } else {
                query = "select id_employee, departmentName as [отдел],jobTitle as [должность], employeeSurname as [фамилия], employeeName as[имя],employeePatronymic as[отчество], bornDate as [дата рождения], passport as [пасспорт],employeeLogin as [логин], employeePassword as [пароль] from employees inner join departments on employees.id_department = departments.id_depatment inner join jobTitles on employees.id_jobTitle = jobTitles.id_jobTitle;";
            }
            SqlDataAdapter adapter = new SqlDataAdapter(query, dataBase.getSrc());
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            dataGridView1.DataSource = dt;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.AutoResizeColumns();
        }
        void ComBx1() {
            string query = "select id_depatment, departmentName from departments;";
            SqlDataAdapter adapter = new SqlDataAdapter(query, dataBase.getSrc());
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            comboBox1.DataSource = dt;
            comboBox1.ValueMember = "id_depatment";
            comboBox1.DisplayMember = "departmentName";
        }
        void ComBx2() {
            string query = "select id_jobTitle, jobTitle from jobTitles;";
            SqlDataAdapter adapter = new SqlDataAdapter(query, dataBase.getSrc());
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            comboBox2.DataSource = dt;
            comboBox2.ValueMember = "id_jobTitle";
            comboBox2.DisplayMember = "jobTitle";
        }
        private void Form5_Load(object sender, EventArgs e) {
        }
        private void textBox1_TextChanged(object sender, EventArgs e) {
            textBox1.Text = System.Text.RegularExpressions.Regex.Replace(textBox1.Text, STRING_PATTERN, "");
        }
        private void textBox3_TextChanged(object sender, EventArgs e) {
            textBox3.Text = System.Text.RegularExpressions.Regex.Replace(textBox3.Text, STRING_PATTERN, "");
        }
        private void textBox4_TextChanged(object sender, EventArgs e) {
            textBox4.Text = System.Text.RegularExpressions.Regex.Replace(textBox4.Text, STRING_PATTERN, "");
        }
        private void button1_Click(object sender, EventArgs e) {
            try {
                string query = "";
                if (textBox1.Text.Length < 5) {
                    MessageBox.Show("Фамилия должна быть больше 4 символов", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (textBox3.Text.Length < 2) {
                    MessageBox.Show("Имя не может быть меньше 2х символов", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (textBox4.Text.Length < 5) {
                    MessageBox.Show("Отчество должно быть больше 4 симвилов", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                DateTime today = DateTime.Today;
                DateTime age;
                DateTime.TryParse(dateTimePicker1.Value.ToString(), out age);
                age = age.AddYears(18);
                if (today < age) {
                    MessageBox.Show("работник должен быть не младше 18 лет", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (int.Parse(comboBox2.SelectedValue.ToString()) == 1) {
                    query = $"Insert into employees values ({comboBox1.SelectedValue}, {comboBox2.SelectedValue},N'{textBox1.Text}', N'{textBox3.Text}', N'{textBox4.Text}', '{dateTimePicker1.Value.ToString("yyyy-MM-dd")}', '{maskedTextBox1.Text}', null, null);";
                    SqlCommand sqlCommand = new SqlCommand(query, sql);
                    sqlCommand.ExecuteNonQuery();
                    MessageBox.Show("Данные успешно добавлены", "SUCCESS", MessageBoxButtons.OK);
                    employees();
                } else {
                    if (textBox5.Text.Length < 4) {
                        MessageBox.Show("логин должен быть больше 3 символов", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    if (textBox6.Text.Length < 3) {
                        MessageBox.Show("пароль должен быть больше 3 символов", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    } else {
                        if (comboBox2.Text == "главный бухгалтер" || comboBox2.Text == "директор" || comboBox2.Text == "зам директора") {
                            SqlDataAdapter adapter2 = new SqlDataAdapter();
                            DataTable table2 = new DataTable();
                            string query2 = $"select jobTitle from employees inner join jobTitles on employees.id_jobTitle = jobTitles.id_jobTitle where jobTitle = N'{comboBox2.Text}'";
                            SqlCommand command2 = new SqlCommand(query2, dataBase.getConnection());
                            adapter2.SelectCommand = command2;
                            adapter2.Fill(table2);
                            if (table2.Rows.Count != 0) {
                                MessageBox.Show("Директор, зам директора и главный бухгалтер может быть только 1", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                            query = $"Insert into employees values ({comboBox1.SelectedValue}, {comboBox2.SelectedValue},N'{textBox1.Text}', N'{textBox3.Text}', N'{textBox4.Text}', '{dateTimePicker1.Value.ToString("yyyy-MM-dd")}', '{maskedTextBox1.Text}',N'{textBox5.Text}', N'{textBox6.Text}');";
                            SqlCommand sqlCommand = new SqlCommand(query, sql);
                            sqlCommand.ExecuteNonQuery();
                            MessageBox.Show("Данные успешно добавлены", "SUCCESS", MessageBoxButtons.OK);
                            employees();
                        } else {
                            query = $"Insert into employees values ({comboBox1.SelectedValue}, {comboBox2.SelectedValue},N'{textBox1.Text}', N'{textBox3.Text}', N'{textBox4.Text}', '{dateTimePicker1.Value.ToString("yyyy-MM-dd")}', '{maskedTextBox1.Text}', N'{textBox5.Text}', N'{textBox6.Text}');";
                            SqlCommand sqlCommand = new SqlCommand(query, sql);
                            sqlCommand.ExecuteNonQuery();
                            MessageBox.Show("Данные успешно добавлены", "SUCCESS", MessageBoxButtons.OK);
                            employees();
                        }
                    }
                }
            } catch (Exception ex) {
                MessageBox.Show($"{ex.Message}", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void button2_Click(object sender, EventArgs e) {
            try {
                if (dataGridView1.Rows.Count == 0) { return; }
                string query = $"Delete From employees Where id_employee = {dataGridView1.CurrentRow.Cells[0].Value}";
                SqlCommand sqlCommand = new SqlCommand(query, sql);
                sqlCommand.ExecuteNonQuery();
                MessageBox.Show("Данные успешно удалены", "SUCCESS", MessageBoxButtons.OK);
                employees();
            } catch (Exception ex) {
                MessageBox.Show($"Не удалось удалить данные, попробуйте удалить данные из связанных таблиц", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e) {
            comboBox1.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            comboBox2.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            textBox1.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            textBox3.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
            textBox4.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
            dateTimePicker1.Value = Convert.ToDateTime(dataGridView1.CurrentRow.Cells[6].Value.ToString());
            maskedTextBox1.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();
            textBox5.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();
            textBox6.Text = dataGridView1.CurrentRow.Cells[9].Value.ToString();
        }
        private void Form5_FormClosing(object sender, FormClosingEventArgs e) {
            sql.Close();
            Form form = Form2.getInstance();
            form.Show();
        }
        private void button3_Click(object sender, EventArgs e) {
            try {
                if (dataGridView1.Rows.Count == 0) { return; }
                if (textBox1.Text.Length < 5) {
                    MessageBox.Show("Фамилия должна быть больше 4 символов", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (textBox3.Text.Length < 2) {
                    MessageBox.Show("Имя не может быть меньше 2х символов", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (textBox4.Text.Length < 5) {
                    MessageBox.Show("Отчество должно быть больше 4 симвилов", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                DateTime today = DateTime.Today;
                DateTime age;
                DateTime.TryParse(dateTimePicker1.Value.ToString(), out age);
                age = age.AddYears(18);
                if (today < age) {
                    MessageBox.Show("работник должен быть не младше 18 лет", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (int.Parse(comboBox2.SelectedValue.ToString()) == 1) {
                    string query = $"Update employees Set employeeSurname = N'{textBox1.Text}', employeeName =  N'{textBox3.Text}', employeePatronymic = N'{textBox4.Text}', bornDate = '{dateTimePicker1.Value.ToString("yyyy-MM-dd")}', passport = '{maskedTextBox1.Text}' ,employeeLogin = null, employeePassword = null  where id_employee = {dataGridView1.CurrentRow.Cells[0].Value}";
                    SqlCommand com = new SqlCommand(query, sql);
                    com.ExecuteNonQuery();
                    MessageBox.Show("Данные успешено изменены", "SUCCESS", MessageBoxButtons.OK);
                    employees();
                } else {
                    if (textBox5.Text.Length < 4) {
                        MessageBox.Show("логин должен быть больше 3 символов", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    if (textBox6.Text.Length < 4) {
                        MessageBox.Show("пароль должен быть больше 3 символов", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    string query = $"Update employees Set employeeSurname = N'{textBox1.Text}', employeeName =  N'{textBox3.Text}', employeePatronymic = N'{textBox4.Text}', bornDate = '{dateTimePicker1.Value.ToString("yyyy-MM-dd")}', passport = '{maskedTextBox1.Text}',employeeLogin = N'{textBox5.Text}', employeePassword = N'{textBox6.Text}'  where id_employee = {dataGridView1.CurrentRow.Cells[0].Value}";
                    SqlCommand com = new SqlCommand(query, sql);
                    com.ExecuteNonQuery();
                    MessageBox.Show("Данные успешено изменены", "SUCCESS", MessageBoxButtons.OK);
                    employees();
                }
            } catch (Exception ex) {
                MessageBox.Show($"{ex.Message}", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void экспортВExcelToolStripMenuItem_Click(object sender, EventArgs e) {
            try {
                Excel.Application exelApp = new Excel.Application();
                exelApp.Workbooks.Add();
                Excel.Worksheet wsh = (Excel.Worksheet)exelApp.ActiveSheet;
                wsh.Rows[1].Style.Font.Size = 12;
                exelApp.Cells[1, 1] = "работники ОАО Медпласт";
                wsh.Range[wsh.Cells[1, 1], wsh.Cells[2, dataGridView1.Rows[0].Cells.Count - 1]].Merge();
                for (int i = 0; i < dataGridView1.RowCount; i++) {
                    for (int j = 1; j < dataGridView1.ColumnCount; j++) {
                        wsh.Columns.AutoFit();
                        wsh.Cells[3, j] = dataGridView1.Columns[j].HeaderText.ToString();
                        wsh.Cells[i + 4, j] = dataGridView1[j, i].Value.ToString();
                    }
                }
                Excel.Range tRange = wsh.UsedRange;
                tRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                tRange.Borders.Weight = Excel.XlBorderWeight.xlThin;
                exelApp.Visible = true;
            } catch { }
        }
        public DataTable GetData(string query) {
            SqlDataAdapter adapter = new SqlDataAdapter(query, dataBase.getSrc());
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            return dt;
        }
        private void radioButton1_CheckedChanged(object sender, EventArgs e) {
            DataTable dt;
            if (radioButton1.Checked == true) {
                button4.Enabled = true;
                comboBox4.Visible = true;
                dt = GetData("select id_depatment, departmentName from departments;");
                comboBox4.DataSource = dt;
                comboBox4.ValueMember = "id_depatment";
                comboBox4.DisplayMember = "departmentName";
            } else {
                comboBox4.Visible = false;
            }
        }
        private void button4_Click(object sender, EventArgs e) {
            if (radioButton1.Checked) {
                if (user.getJobTitle() != "администратор") {
                    filter = $"select id_employee, departmentName as [отдел],jobTitle as [должность], employeeSurname as [фамилия], employeeName as[имя],employeePatronymic as[отчество], bornDate as [дата рождения] from employees inner join departments on employees.id_department = departments.id_depatment inner join jobTitles on employees.id_jobTitle = jobTitles.id_jobTitle where departments.id_depatment = {comboBox4.SelectedValue};";
                } else {
                    filter = $"select id_employee, departmentName as [отдел],jobTitle as [должность], employeeSurname as [фамилия], employeeName as[имя],employeePatronymic as[отчество], bornDate as [дата рождения], employeeLogin as [логин], employeePassword as [пароль] from employees inner join departments on employees.id_department = departments.id_depatment inner join jobTitles on employees.id_jobTitle = jobTitles.id_jobTitle where departments.id_depatment = {comboBox4.SelectedValue};";
                }
            }
            if (radioButton2.Checked) {
                if (user.getJobTitle() != "администратор") {
                    filter = $"select id_employee, departmentName as [отдел],jobTitle as [должность], employeeSurname as [фамилия], employeeName as[имя],employeePatronymic as[отчество], bornDate as [дата рождения] from employees inner join departments on employees.id_department = departments.id_depatment inner join jobTitles on employees.id_jobTitle = jobTitles.id_jobTitle where jobTitles.id_jobTitle = {comboBox4.SelectedValue};";
                } else {
                    filter = $"select id_employee, departmentName as [отдел],jobTitle as [должность], employeeSurname as [фамилия], employeeName as[имя],employeePatronymic as[отчество], bornDate as [дата рождения], employeeLogin as [логин], employeePassword as [пароль] from employees inner join departments on employees.id_department = departments.id_depatment inner join jobTitles on employees.id_jobTitle = jobTitles.id_jobTitle where jobTitles.id_jobTitle = {comboBox4.SelectedValue};";
                }
            }
            if (radioButton3.Checked) {
                if (textBox7.Text.Length == 0) {
                    MessageBox.Show($"Для фильтрации необходимо заполнить данные", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                } else {
                    if (user.getJobTitle() != "администратор") {
                        filter = $"select id_employee, departmentName as [отдел],jobTitle as [должность], employeeSurname as [фамилия], employeeName as[имя],employeePatronymic as[отчество], bornDate as [дата рождения] from employees inner join departments on employees.id_department = departments.id_depatment inner join jobTitles on employees.id_jobTitle = jobTitles.id_jobTitle where employeeSurname like N'{textBox7.Text}';";
                    } else {
                        filter = $"select id_employee, departmentName as [отдел],jobTitle as [должность], employeeSurname as [фамилия], employeeName as[имя],employeePatronymic as[отчество], bornDate as [дата рождения], employeeLogin as [логин], employeePassword as [пароль] from employees inner join departments on employees.id_department = departments.id_depatment inner join jobTitles on employees.id_jobTitle = jobTitles.id_jobTitle where employeeSurname like N'{textBox7.Text}';";
                    }
                }
            }
            if (radioButton4.Checked) {
                if (user.getJobTitle() != "администратор") {
                    filter = $"select id_employee, departmentName as [отдел],jobTitle as [должность], employeeSurname as [фамилия], employeeName as[имя],employeePatronymic as[отчество], bornDate as [дата рождения] from employees inner join departments on employees.id_department = departments.id_depatment inner join jobTitles on employees.id_jobTitle = jobTitles.id_jobTitle where bornDate >= '{dateTimePicker2.Value.ToString("yyyy - MM - dd")}' and bornDate <= '{dateTimePicker3.Value.ToString("yyyy-MM-dd")}';";
                } else {
                    filter = $"select id_employee, departmentName as [отдел],jobTitle as [должность], employeeSurname as [фамилия], employeeName as[имя],employeePatronymic as[отчество], bornDate as [дата рождения], employeeLogin as [логин], employeePassword as [пароль] from employees inner join departments on employees.id_department = departments.id_depatment inner join jobTitles on employees.id_jobTitle = jobTitles.id_jobTitle where bornDate >= '{dateTimePicker2.Value.ToString("yyyy-MM-dd")}' and bornDate <= '{dateTimePicker3.Value.ToString("yyyy-MM-dd")}';";
                }
            }
            DataTable dt = GetData(filter);
            dataGridView1.DataSource = dt;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.AutoResizeColumns();
            button6.Enabled = true;
        }
        private void radioButton2_CheckedChanged(object sender, EventArgs e) {
            DataTable dt;
            if (radioButton2.Checked == true) {
                button4.Enabled = true;
                comboBox4.Visible = true;
                dt = GetData("select id_jobTitle, jobTitle from jobTitles;");
                comboBox4.DataSource = dt;
                comboBox4.ValueMember = "id_jobTitle";
                comboBox4.DisplayMember = "jobTitle";
            } else {
                comboBox4.Visible = false;
            }
        }
        private void radioButton3_CheckedChanged(object sender, EventArgs e) {
            if (radioButton3.Checked == true) {
                button4.Enabled = true;
                textBox7.Visible = true;
            } else {
                textBox7.Visible = false;
            }
        }
        private void radioButton4_CheckedChanged(object sender, EventArgs e) {
            if (radioButton4.Checked == true) {
                button4.Enabled = true;
                dateTimePicker2.Visible = true;
                dateTimePicker3.Visible = true;
            } else {
                dateTimePicker2.Visible = false;
                dateTimePicker3.Visible = false;
            }
        }
        private void textBox7_TextChanged(object sender, EventArgs e) {
            textBox7.Text = System.Text.RegularExpressions.Regex.Replace(textBox7.Text, STRING_PATTERN, "");
        }
        private void button6_Click(object sender, EventArgs e) {
            employees();
            button6.Enabled = false;
        }
        private void Form5_Shown(object sender, EventArgs e) {
            sql = dataBase.getConnection();
            sql.Open();
            setMode();
            getAccess();
            dateTimePicker1.MaxDate = DateTime.Now;
            dateTimePicker1.MinDate = new DateTime(1950, 01, 01);
            employees();
            ComBx1();
            ComBx2();
        }

        private void label11_Click(object sender, EventArgs e) {

        }
    }
}
