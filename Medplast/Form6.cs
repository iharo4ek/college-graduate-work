using System;using System.Collections.Generic;using System.ComponentModel;using System.Data;using System.Data.SqlClient;using System.Drawing;using System.Linq;using System.Text;using System.Threading.Tasks;using System.Windows.Forms;using Excel = Microsoft.Office.Interop.Excel;namespace Medplast {    public partial class Form6 : Form {        private Modes mode;        private string filter = "";        private const string STRING_PATTERN = @"[^а-яА-яa-zA-Z]";        private const string INT_PATTERN = @"[^0-9]";        DataBase dataBase = DataBase.getInstance();        private static Form instance;        private SqlConnection sql;        private Form6() {            InitializeComponent();        }        public static Form getInstance() {            if (instance == null) {                instance = new Form6();            }            return instance;        }        private void cars() {            string query = "select id_car, numberOfTheCar as  [гос. номер],(employeeSurname + ' ' + employeeName + ' ' + employeePatronymic) as [водитель],carBrand as [марка], dateOfpurchase as [дата покупки] from cars inner join employees on cars.id_driver = employees.id_employee;";            SqlDataAdapter adapter = new SqlDataAdapter(query, dataBase.getSrc());            DataTable dt = new DataTable();            adapter.Fill(dt);            dataGridView1.DataSource = dt;            dataGridView1.Columns[0].Visible = false;            dataGridView1.AutoResizeColumns();        }        private void setMode() {            User user = User.getInstance();            switch (user.getJobTitle()) {                case "директор": {                        this.mode = Modes.READ;                        break;                    }                case "зам директора": {                        this.mode = Modes.READ;                        break;                    }                case "главный бухгалтер": {                        this.mode = Modes.READ;                        break;                    }                case "бухгалтер": {                        this.mode = Modes.READ;                        break;                    }
                default: {
                        this.mode = Modes.READWRITE;
                        break;                    }            }        }        private void getAccess() {            if (mode == Modes.READ) {                button1.Visible = false;                button2.Visible = false;                button3.Visible = false;                textBox1.Visible = false;                maskedTextBox1.Visible = false;                dateTimePicker1.Visible = false;                label1.Visible = false;                label2.Visible = false;                label3.Visible = false;                label5.Visible = false;                comboBox1.Visible = false;            }        }        void ComBx1() {
            string query = "select id_employee,(employeeSurname + ' ' + employeeName + ' ' + employeePatronymic) as [driver] from employees;";
            SqlDataAdapter adapter = new SqlDataAdapter(query, dataBase.getSrc());
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            comboBox1.DataSource = dt;
            comboBox1.ValueMember = "id_employee";
            comboBox1.DisplayMember = "driver";
        }        private void Form6_Load(object sender, EventArgs e) {        }        private void Form6_FormClosing(object sender, FormClosingEventArgs e) {            sql.Close();            Form form = Form2.getInstance();            form.Show();        }        private void button5_Click(object sender, EventArgs e) {            for (int i = 0; i < dataGridView1.RowCount; i++) {                dataGridView1.Rows[i].Selected = false;                for (int j = 0; j < dataGridView1.ColumnCount; j++)                    if (dataGridView1.Rows[i].Cells[j].Value != null)                        if (dataGridView1.Rows[i].Cells[j].Value.ToString().ToLower().Contains(textBox2.Text.ToLower())) {                            dataGridView1.Rows[i].Selected = true;                            break;                        }            }        }        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e) {            maskedTextBox1.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();            comboBox1.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            textBox1.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();            dateTimePicker1.Value = Convert.ToDateTime(dataGridView1.CurrentRow.Cells[4].Value.ToString());        }        private void button2_Click(object sender, EventArgs e) {            try {                if (dataGridView1.Rows.Count == 0) { return; }                string query = $"Delete From cars Where id_car = {dataGridView1.CurrentRow.Cells[0].Value}";                SqlCommand sqlCommand = new SqlCommand(query, sql);                sqlCommand.ExecuteNonQuery();                MessageBox.Show("Данные успешно удалены", "SUCCESS", MessageBoxButtons.OK);                cars();            } catch (Exception ex) {                MessageBox.Show($"Не удалось удалить данные, попробуйте удалить данные из связанных таблиц", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);            }        }        private void button1_Click(object sender, EventArgs e) {            try {                if (maskedTextBox1.Text.Length < 9) {                    MessageBox.Show("Пожылуйста заполните поле гос. номер", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);                    return;                }
                if (textBox1.Text.Length < 3) {
                    MessageBox.Show("Слишком короткое название марки", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                string query = $"Insert into cars values (N'{maskedTextBox1.Text}', {comboBox1.SelectedValue}, N'{textBox1.Text}', '{dateTimePicker1.Value.ToString("yyyy-MM-dd")}');";
                SqlCommand sqlCommand = new SqlCommand(query, sql);
                sqlCommand.ExecuteNonQuery();
                MessageBox.Show("Данные успешно добавлены", "SUCCESS", MessageBoxButtons.OK);
                cars();
            } catch (Exception ex) {                MessageBox.Show($"{ex.Message}", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);            }        }        private void button3_Click(object sender, EventArgs e) {            try {                if (dataGridView1.Rows.Count == 0) { return; }                if (maskedTextBox1.Text.Length < 9) {                    MessageBox.Show("Пожалуйста заполните поле гос. номер", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);                    return;                }
                if (textBox1.Text.Length < 3) {
                    MessageBox.Show("Слишком короткое название марки", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);                    return;
                }
                string query = $"Update cars Set numberOfTheCar = N'{maskedTextBox1.Text}', id_driver = {comboBox1.SelectedValue},carBrand = N'{textBox1.Text}', dateOfpurchase = '{dateTimePicker1.Value.ToString("yyyy-MM-dd")}'  where id_car = {dataGridView1.CurrentRow.Cells[0].Value}";
                SqlCommand com = new SqlCommand(query, sql);
                com.ExecuteNonQuery();
                cars();
                MessageBox.Show("Данные успешно изменены", "SUCCESS", MessageBoxButtons.OK);            } catch (Exception ex) {                MessageBox.Show($"Не удалось изменить данные Произошла ошибка", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);            }
        }        private void textBox1_TextChanged(object sender, EventArgs e) {            textBox1.Text = System.Text.RegularExpressions.Regex.Replace(textBox1.Text, STRING_PATTERN, "");        }        private void экспортВExcelToolStripMenuItem_Click(object sender, EventArgs e) {            try {                Excel.Application exelApp = new Excel.Application();                exelApp.Workbooks.Add();                Excel.Worksheet wsh = (Excel.Worksheet)exelApp.ActiveSheet;                wsh.Rows[1].Style.Font.Size = 12;
                exelApp.Cells[1, 1] = "машины ОАО Медпласт";
                for (int i = 0; i < dataGridView1.RowCount; i++) {                    for (int j = 1; j < dataGridView1.ColumnCount; j++) {                        wsh.Columns.AutoFit();                        wsh.Cells[3, j] = dataGridView1.Columns[j].HeaderText.ToString();                        wsh.Cells[i + 4, j] = dataGridView1[j, i].Value.ToString();                    }                }                Excel.Range tRange = wsh.UsedRange;                tRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;                tRange.Borders.Weight = Excel.XlBorderWeight.xlThin;                exelApp.Visible = true;            } catch { }        }        public DataTable GetData(string query) {            SqlDataAdapter adapter = new SqlDataAdapter(query, dataBase.getSrc());            DataTable dt = new DataTable();            adapter.Fill(dt);            return dt;        }
        private void radioButton2_CheckedChanged(object sender, EventArgs e) {
            if (radioButton2.Checked == true) {
                dateTimePicker2.Visible = true;
                dateTimePicker3.Visible = true;
                button4.Enabled = true;
            } else {
                dateTimePicker2.Visible = false;
                dateTimePicker3.Visible = false;
            }
        }
        private void button4_Click(object sender, EventArgs e) {
            if (radioButton1.Checked) {
                if (textBox3.Text.Length == 0) {
                    MessageBox.Show($"Для фильтрации необходимо заполнить данные", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                } else {
                    filter = $"select id_car, numberOfTheCar as  [гос. номер],(employeeSurname + ' ' + employeeName + ' ' + employeePatronymic) as [водитель],carBrand as [марка], dateOfpurchase as [дата покупки] from cars inner join employees on cars.id_driver = employees.id_employee where carBrand like N'{textBox3.Text}';";
                }
            }
            if (radioButton2.Checked) {
                filter = $"select id_car, numberOfTheCar as  [гос. номер],(employeeSurname + ' ' + employeeName + ' ' + employeePatronymic) as [водитель],carBrand as [марка], dateOfpurchase as [дата покупки] from cars inner join employees on cars.id_driver = employees.id_employee where dateOfpurchase >= '{dateTimePicker2.Value.ToString("yyyy-MM-dd")}' and  dateOfpurchase <= '{dateTimePicker3.Value.ToString("yyyy-MM-dd")}';";
            }
            DataTable dt = GetData(filter);
            dataGridView1.DataSource = dt;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.AutoResizeColumns();
            button6.Enabled = true;
        }
        private void textBox3_TextChanged(object sender, EventArgs e) {
        }
        private void radioButton1_CheckedChanged(object sender, EventArgs e) {
            if (radioButton1.Checked == true) {
                textBox3.Visible = true;
                button4.Enabled = true;
            } else {
                textBox3.Visible = false;
            }
        }
        private void button6_Click(object sender, EventArgs e) {
            cars();
            button6.Enabled = false;
        }
        private void Form6_Shown(object sender, EventArgs e) {
            sql = dataBase.getConnection();            sql.Open();            setMode();            getAccess();            cars();            ComBx1();            dateTimePicker1.MaxDate = DateTime.Now;            dateTimePicker1.MinDate = new DateTime(1900, 01, 01);
        }

        private void maskedTextBox1_MaskInputRejected(object sender, MaskInputRejectedEventArgs e) {

        }
    }}