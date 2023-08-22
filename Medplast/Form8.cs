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
    public partial class Form8 : Form {
        private string filter = "";
        private Modes mode;
        private const string STRING_PATTERN = @"[^а-яА-яa-zA-Z]";
        private const string INT_PATTERN = @"[^0-9]";
        DataBase dataBase = DataBase.getInstance();
        private SqlConnection sql;
        private Form8() {
            InitializeComponent();
        }
        private static Form instance;
        public static Form getInstance() {
            if (instance == null) {
                instance = new Form8();
            }
            return instance;
        }
        private void providers() {
            string query = "select id_provider, nameOrganisation as [название организации], adress as [адрес], phoneNumber as [номер телефона], YNP as [УНП] from providers;";
            SqlDataAdapter adapter = new SqlDataAdapter(query, dataBase.getSrc());
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            dataGridView1.DataSource = dt;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.AutoResizeColumns();
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
                maskedTextBox1.Visible = false;
                maskedTextBox2.Visible = false;
                label4.Visible = false;
                label2.Visible = false;
                label1.Visible = false;
                label3.Visible = false;
            }
        }
        private void Form8_Load(object sender, EventArgs e) {
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
        private void button1_Click(object sender, EventArgs e) {
            try {
                DataTable dataTable = GetData($"select phoneNumber from clients where phoneNumber = '{maskedTextBox1.Text}';");
                if (dataTable.Rows.Count != 0) {
                    MessageBox.Show("Данный номер телефона уже зарегестрирован на клиента", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                dataTable = GetData($"select YNP from clients where YNP = '{maskedTextBox2.Text}';");
                if (dataTable.Rows.Count != 0) {
                    MessageBox.Show("Данный УНП уже зарегестрирован на клинета", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (textBox1.Text.Length < 6) {
                    MessageBox.Show("Слишком короткое название организации", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (textBox3.Text.Length < 15) {
                    MessageBox.Show("Слишком короткий адрес", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (maskedTextBox1.Text.Length < 16) {
                    MessageBox.Show("Пожалуйста заполните поле номер телефона", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (maskedTextBox2.Text.Length < 9) {
                    MessageBox.Show("Пожалуйста заполните поле унп", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                string query = $"Insert into providers values (N'{textBox1.Text}', N'{textBox3.Text}', N'{maskedTextBox1.Text}', '{maskedTextBox2.Text}');";
                SqlCommand sqlCommand = new SqlCommand(query, sql);
                sqlCommand.ExecuteNonQuery();
                MessageBox.Show("Данные успешно добавлены", "SUCCESS", MessageBoxButtons.OK);
                providers();
            } catch (Exception ex) {
                MessageBox.Show($"{ex.Message}", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void button2_Click(object sender, EventArgs e) {
            try {
                if (dataGridView1.Rows.Count == 0) { return; }
                string query = $"Delete From providers Where id_provider = {dataGridView1.CurrentRow.Cells[0].Value}";
                SqlCommand sqlCommand = new SqlCommand(query, sql);
                sqlCommand.ExecuteNonQuery();
                MessageBox.Show("Данные успешно удалены", "SUCCESS", MessageBoxButtons.OK);
                providers();
            } catch (Exception ex) {
                MessageBox.Show($"Не удалось удалить данные, попробуйте удалить данные из связанных таблиц", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void button3_Click(object sender, EventArgs e) {
            try {
                DataTable dataTable = GetData($"select phoneNumber from clients where phoneNumber = '{maskedTextBox1.Text}';");
                if (dataTable.Rows.Count != 0) {
                    MessageBox.Show("Данный номер телефона уже зарегестрирован на клиента", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                dataTable = GetData($"select YNP from clients where YNP = '{maskedTextBox2.Text}';");
                if (dataTable.Rows.Count != 0) {
                    MessageBox.Show("Данный УНП уже зарегестрирован на клинета", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (dataGridView1.Rows.Count == 0) { return; }
                if (textBox1.Text.Length < 6) {
                    MessageBox.Show("Слишком короткое название организации", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (textBox3.Text.Length < 15) {
                    MessageBox.Show("Слишком короткий адрес", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (maskedTextBox1.Text.Length < 16) {
                    MessageBox.Show("Пожалуйста заполните поле номер телефона", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                string query = $"Update providers Set nameOrganisation = N'{textBox1.Text}', adress = N'{textBox3.Text}', phoneNumber = N'{maskedTextBox1.Text}', YNP = '{maskedTextBox2.Text}'  where id_provider = {dataGridView1.CurrentRow.Cells[0].Value}";
                SqlCommand com = new SqlCommand(query, sql);
                com.ExecuteNonQuery();
                providers();
                MessageBox.Show("Данные успешно изменены", "SUCCESS", MessageBoxButtons.OK);
            } catch (Exception ex) {
                MessageBox.Show($"Не удалось изменить данные", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void Form8_FormClosing(object sender, FormClosingEventArgs e) {
            sql.Close();
            Form form = Form2.getInstance();
            form.Show();
        }
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e) {
            textBox1.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            textBox3.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            maskedTextBox1.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            maskedTextBox2.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
        }
        private void textBox1_TextChanged(object sender, EventArgs e) {
            //textBox1.Text = System.Text.RegularExpressions.Regex.Replace(textBox1.Text, STRING_PATTERN, "");
        }
        private void textBox3_TextChanged(object sender, EventArgs e) {
        }
        private void экспортВExcelToolStripMenuItem_Click(object sender, EventArgs e) {
            try {
                Excel.Application exelApp = new Excel.Application();
                exelApp.Workbooks.Add();
                Excel.Worksheet wsh = (Excel.Worksheet)exelApp.ActiveSheet;
                wsh.Rows[1].Style.Font.Size = 12;
                exelApp.Cells[1, 1] = "список поставщиков ОАО Медпласт";
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
            if (radioButton1.Checked == true) {
                maskedTextBox3.Visible = false;
                button4.Enabled = true;
                textBox4.Text = "";
                textBox4.Visible = true;
            }
        }
        private void radioButton2_CheckedChanged(object sender, EventArgs e) {
            if (radioButton2.Checked == true) {
                maskedTextBox3.Visible = false;
                button4.Enabled = true;
                textBox4.Text = "";
                textBox4.Visible = true;
            }
        }
        private void radioButton3_CheckedChanged(object sender, EventArgs e) {
            if (radioButton3.Checked == true) {
                textBox4.Visible = false;
                button4.Enabled = true;
                maskedTextBox3.Text = "";
                maskedTextBox3.Mask = maskedTextBox1.Mask;
                maskedTextBox3.Visible = true;
            }
        }
        private void radioButton4_CheckedChanged(object sender, EventArgs e) {
            if (radioButton4.Checked == true) {
                textBox4.Visible = false;
                button4.Enabled = true;
                maskedTextBox3.Text = "";
                maskedTextBox3.Mask = maskedTextBox2.Mask;
                maskedTextBox3.Visible = true;
            }
        }
        private void button4_Click(object sender, EventArgs e) {
            if (radioButton1.Checked) {
                if (textBox4.Text.Length == 0) {
                    MessageBox.Show($"Для фильтрации необходимо заполнить данные", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                } else {
                    filter = $"select id_provider, nameOrganisation as [название организации], adress as [адрес], phoneNumber as [номер телефона], YNP as [УНП] from providers where nameOrganisation like N'{textBox4.Text}';";
                }
            }
            if (radioButton2.Checked) {
                if (textBox4.Text.Length == 0) {
                    MessageBox.Show($"Для фильтрации необходимо заполнить данные", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                } else {
                    filter = $"select id_provider, nameOrganisation as [название организации], adress as [адрес], phoneNumber as [номер телефона], YNP as [УНП] from providers where adress like N'{textBox4.Text}';";
                }
            }
            if (radioButton3.Checked) {
                if (maskedTextBox3.Text.Length < 16) {
                    MessageBox.Show($"Для фильтрации необходимо заполнить данные", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                } else {
                    filter = $"select id_provider, nameOrganisation as [название организации], adress as [адрес], phoneNumber as [номер телефона], YNP as [УНП] from providers where phoneNumber like N'{maskedTextBox3.Text}';";
                }
            }
            if (radioButton4.Checked) {
                if (maskedTextBox3.Text.Length < 9) {
                    MessageBox.Show($"Для фильтрации необходимо заполнить данные", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                } else {
                    filter = $"select id_provider, nameOrganisation as [название организации], adress as [адрес], phoneNumber as [номер телефона], YNP as [УНП] from providers where YNP like N'{maskedTextBox3.Text}';";
                }
            }
            DataTable dt = GetData(filter);
            dataGridView1.DataSource = dt;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.AutoResizeColumns();
            button6.Enabled = true;
        }
        private void button6_Click(object sender, EventArgs e) {
            providers();
            button6.Enabled = false;
        }
        private void Form8_Shown(object sender, EventArgs e) {
            sql = dataBase.getConnection();
            sql.Open();
            setMode();
            getAccess();
            providers();
        }
        private void maskedTextBox3_MaskInputRejected(object sender, MaskInputRejectedEventArgs e) {

        }

        private void экспортВExcelToolStripMenuItem_MouseHover(object sender, EventArgs e) {
            this.BackColor = Color.Teal;
        }
    }
}
