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
    public partial class Form4 : Form {
        private static Form4 instance;
        DataBase dataBase = DataBase.getInstance();
        private Modes mode;
        private const string STRING_PATTERN = @"[^а-яА-яa-zA-Z ]";
        private const string INT_PATTERN = @"[^0-9]";
        private SqlConnection sql;
        private Form4() {
            InitializeComponent();
        }
        public static Form4 getInstance() {
            if (instance == null)
                instance = new Form4();
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
                label1.Visible = false;
            }
        }
        private void Form4_Load(object sender, EventArgs e) {
        }
        private void button5_Click(object sender, EventArgs e) {
            for (int i = 0; i < dataGridView1.RowCount; i++) {
                dataGridView1.Rows[i].Selected = false;
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    if (dataGridView1.Rows[i].Cells[j].Value != null)
                        if (dataGridView1.Rows[i].Cells[j].Value.ToString().ToLower().Contains(textBox2.Text.ToLower())) {
                            dataGridView1.Rows[i].Selected = true;
                            break;
                        }
            }
        }
        private void departments() {
            string query = "select id_depatment, departmentName as [отдел], countEmployess as [количество сотрудников] from departments";
            SqlDataAdapter adapter = new SqlDataAdapter(query, dataBase.getSrc());
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            dataGridView1.DataSource = dt;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.AutoResizeColumns();
        }
        private void button1_Click(object sender, EventArgs e) {
            try {
                if (textBox1.Text.Length == 0) {
                    MessageBox.Show("Сначала необходимо заполнить поле название отдела", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (textBox1.Text.Length < 6) {
                    MessageBox.Show("Слишком короткое название отдела", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                string query = $"Insert into departments values (N'{textBox1.Text}', 0);";
                SqlCommand sqlCommand = new SqlCommand(query, sql);
                sqlCommand.ExecuteNonQuery();
                MessageBox.Show("Данные успешно добавлены", "SUCCESS", MessageBoxButtons.OK);
                departments();
            } catch (Exception ex) {
                MessageBox.Show($"{ex.Message}", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void textBox1_TextChanged(object sender, EventArgs e) {
            textBox1.Text = System.Text.RegularExpressions.Regex.Replace(textBox1.Text, STRING_PATTERN, "");
        }
        private void button2_Click(object sender, EventArgs e) {
            try {
                if (dataGridView1.Rows.Count == 0) { return; }
                string query = $"Delete From departments Where id_depatment = {dataGridView1.CurrentRow.Cells[0].Value}";
                SqlCommand sqlCommand = new SqlCommand(query, sql);
                sqlCommand.ExecuteNonQuery();
                MessageBox.Show("Данные успешно удалены", "SUCCESS", MessageBoxButtons.OK);
                departments();
            } catch (Exception ex) {
                MessageBox.Show($"Не удалось удалить данные, попробуйте удалить данные из связанных таблиц", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void button3_Click(object sender, EventArgs e) {
            try {
                if (dataGridView1.Rows.Count == 0) { return; }
                if (textBox1.Text.Length < 6) {
                    MessageBox.Show("Слишком короткое название отдела", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                string query = $"Update departments Set departmentName = N'{textBox1.Text}'  where id_depatment = {dataGridView1.CurrentRow.Cells[0].Value}";
                SqlCommand com = new SqlCommand(query, sql);
                com.ExecuteNonQuery();
                departments();
                MessageBox.Show("Данные успешно изменены", "SUCCESS", MessageBoxButtons.OK);
            } catch (Exception ex) {
                MessageBox.Show($"Не удалось изменить данные", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e) {
            textBox1.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
        }
        private void Form4_FormClosing(object sender, FormClosingEventArgs e) {
            Form form = Form2.getInstance();
            sql.Close();
            form.Show();
        }
        private void экспортВExcelToolStripMenuItem_Click(object sender, EventArgs e) {
            try {
                Excel.Application exelApp = new Excel.Application();
                exelApp.Workbooks.Add();
                Excel.Worksheet wsh = (Excel.Worksheet)exelApp.ActiveSheet;
                wsh.Rows[1].Style.Font.Size = 12;
                exelApp.Cells[1, 1] = "отделы ОАО Медпласт";
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
        private void Form4_Shown(object sender, EventArgs e) {
            sql = dataBase.getConnection();
            sql.Open();
            departments();
            setMode();
            getAccess();
        }
    }
}
