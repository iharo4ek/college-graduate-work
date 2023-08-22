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
using Word = Microsoft.Office.Interop.Word;
namespace Medplast {
    public partial class Form11 : Form {
        private Modes mode;
        private const string STRING_PATTERN = @"[^а-яА-яa-zA-Z ]";
        private const string INT_PATTERN = @"[^0-9]";
        DataBase dataBase = DataBase.getInstance();
        User user = User.getInstance();
        private SqlConnection sql;
        private string filter = "";
        private Form11() {
            InitializeComponent();
        }
        private static Form instance;
        public static Form getInstance() {
            if (instance == null) {
                instance = new Form11();
            }
            return instance;
        }
        private void materials() {
            string query = "select id_material, materialName as [название материала], unit as [ед. изм.], round(countAtStore,2) as [количество на складе] from materials";
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
                case "инженер": {
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
                label4.Visible = false;
                label1.Visible = false;
                comboBox1.Visible = false;
            }
        }
        private void Form11_Load(object sender, EventArgs e) {
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
        private void textBox1_TextChanged(object sender, EventArgs e) {
            textBox1.Text = System.Text.RegularExpressions.Regex.Replace(textBox1.Text, STRING_PATTERN, "");
        }
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e) {
            textBox1.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            comboBox1.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
        }
        private void Form11_FormClosing(object sender, FormClosingEventArgs e) {
            sql.Close();
            Form form = Form2.getInstance();
            form.Show();
        }
        private void button1_Click(object sender, EventArgs e) {
            try {
                if (textBox1.Text.Length < 5) {
                    MessageBox.Show("Слишком короткое название материала", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (comboBox1.Text == "") {
                    MessageBox.Show("Выберите единицу измерения", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                string query = $"Insert into materials values (N'{textBox1.Text}',N'{comboBox1.Text}', 0);";
                SqlCommand sqlCommand = new SqlCommand(query, sql);
                sqlCommand.ExecuteNonQuery();
                MessageBox.Show("Данные успешно добавлены", "SUCCESS", MessageBoxButtons.OK);
                materials();
            } catch (Exception ex) {
                MessageBox.Show($"{ex.Message}", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void button2_Click(object sender, EventArgs e) {
            try {
                if (dataGridView1.Rows.Count == 0) { return; }
                string query = $"Delete From materials Where id_material = {dataGridView1.CurrentRow.Cells[0].Value}";
                SqlCommand sqlCommand = new SqlCommand(query, sql);
                sqlCommand.ExecuteNonQuery();
                MessageBox.Show("Данные успешно удалены", "SUCCESS", MessageBoxButtons.OK);
                materials();
            } catch (Exception ex) {
                MessageBox.Show($"Не удалось удалить данные, попробуйте удалить данные из связанных таблиц", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void button3_Click(object sender, EventArgs e) {
            try {
                if (dataGridView1.Rows.Count == 0) { return; }
                if (textBox1.Text.Length < 6) {
                    MessageBox.Show("Слишком короткое название должности", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                string query = $"Update materials Set materialName = N'{textBox1.Text}', unit = N'{comboBox1.Text}' where id_material = {dataGridView1.CurrentRow.Cells[0].Value}";
                SqlCommand com = new SqlCommand(query, sql);
                com.ExecuteNonQuery();
                materials();
                MessageBox.Show("Данные успешно изменены", "SUCCESS", MessageBoxButtons.OK);
            } catch (Exception ex) {
                MessageBox.Show($"Не удалось изменить данные", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void экспортВExcelToolStripMenuItem_Click(object sender, EventArgs e) {
            try {
                Excel.Application exelApp = new Excel.Application();
                exelApp.Workbooks.Add();
                Excel.Worksheet wsh = (Excel.Worksheet)exelApp.ActiveSheet;
                wsh.Rows[1].Style.Font.Size = 12;
                exelApp.Cells[1, 1] = "список материалов, используемых в производстве ОАО Медпласт";
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
        private DataTable GetData(string query) {
            SqlDataAdapter adapter = new SqlDataAdapter(query, dataBase.getSrc());
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            return dt;
        }
        private void ReplaceWordStub(string stubToReplace, string text, Word.Document wordDocumet) {
            var range = wordDocumet.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: stubToReplace, ReplaceWith: text);
        }
        private void сформироватьОстаткиПоскладуToolStripMenuItem_Click(object sender, EventArgs e) {
            if (dataGridView1.Rows.Count == 0) {
                MessageBox.Show("Произошла ошибка", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            var wordapp = new Word.Application();
            string path = Environment.CurrentDirectory + @"\Sclad.docx";
            var wordDocument = wordapp.Documents.Open(path);
            try {
                int ind = 1;

                wordapp.Visible = false;
                DataTable dt = GetData($"select productName, round(cost,2), countAtStore from products");
                Word.Table tb = wordDocument.Tables[1];
                ReplaceWordStub("{date}", DateTime.Now.ToString(), wordDocument);
                int count = 0; double sum = 0;
                Word.Row r = tb.Rows[2];
                for (int index = 0; index < dt.Rows.Count; index++, ind++) {
                    count += int.Parse(dt.Rows[index].ItemArray[2].ToString());
                    sum += int.Parse(dt.Rows[index].ItemArray[2].ToString()) * double.Parse(dt.Rows[index].ItemArray[1].ToString());
                    r.Cells[1].Range.Text = ind.ToString();
                    r.Cells[2].Range.Text = dt.Rows[index].ItemArray[0].ToString();
                    r.Cells[3].Range.Text = "шт";
                    r.Cells[4].Range.Text = dt.Rows[index].ItemArray[2].ToString();
                    r.Cells[5].Range.Text = dt.Rows[index].ItemArray[1].ToString();
                    r.Cells[6].Range.Text = (int.Parse(dt.Rows[index].ItemArray[2].ToString()) * double.Parse(dt.Rows[index].ItemArray[1].ToString())).ToString();
                    r = tb.Rows.Add();
                }
                wordDocument.Range(r.Cells[1].Range.Start, r.Cells[2].Range.End).Cells.Merge();
                r.Cells[1].Range.Text = "ИТОГО";
                r.Cells[2].Range.Text = "{obshC}";
                r.Cells[4].Range.Text = "{obshS}";
                ReplaceWordStub("{obshC}", count.ToString(), wordDocument);
                ReplaceWordStub("{obshS}", sum.ToString(), wordDocument);
                ReplaceWordStub("{date}", DateTime.Now.ToString("dd.MM.yyyy"), wordDocument);
                tb = wordDocument.Tables[2];
                double count2 = 0; ind = 1;
                dt = GetData("select materialName, countAtStore, unit from materials");
                r = tb.Rows[2];
                for (int index = 0; index < dt.Rows.Count; index++, ind++) {
                    r.Cells[1].Range.Text = ind.ToString();
                    r.Cells[2].Range.Text = dt.Rows[index].ItemArray[0].ToString();
                    r.Cells[3].Range.Text = dt.Rows[index].ItemArray[2].ToString();
                    r.Cells[4].Range.Text = dt.Rows[index].ItemArray[1].ToString();
                    count2 += double.Parse(dt.Rows[index].ItemArray[1].ToString());
                    r = tb.Rows.Add();
                }
                wordDocument.Range(r.Cells[1].Range.Start, r.Cells[2].Range.End).Cells.Merge();
                r.Cells[1].Range.Text = "ИТОГО";
                r.Cells[2].Range.Text = "{jbshC2}";
                ReplaceWordStub("{jbshC2}", count2.ToString(), wordDocument);
                string us = user.getSName() + " " + user.getName() + " " + user.getP();
                ReplaceWordStub("{empl}", us, wordDocument);
                wordapp.Visible = true;
            } catch (Exception ex) {
                MessageBox.Show($"Произошла ошибка", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (wordDocument == null) { return; }
                wordDocument.Close();
            }
        }
        private void radioButton1_CheckedChanged(object sender, EventArgs e) {
            if (radioButton1.Checked == true) {
                button4.Enabled = true;
                numericUpDown1.Visible = false;
                textBox3.Text = "";
                textBox3.Visible = true;
            }
        }
        private void radioButton2_CheckedChanged(object sender, EventArgs e) {
            if (radioButton2.Checked == true) {
                button4.Enabled = true;
                textBox3.Visible = false;
                numericUpDown1.Visible = true;
            }
        }
        private void button4_Click(object sender, EventArgs e) {
            if (radioButton1.Checked == true) {
                if (textBox3.Text.Length <= 0) {
                    MessageBox.Show($"Для фильтрации необходимо заполнить данные", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                } else {
                    filter = $"select id_material, materialName as [название материала], unit as [ед. изм.], countAtStore as [количество на складе] from materials where materialName = N'{textBox3.Text}'";
                }
            }
            if (radioButton2.Checked == true) {
                if (numericUpDown1.Value < 0) {
                    MessageBox.Show($"Неверные данные фильтрации", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                } else {
                    filter = $"select id_material, materialName as [название материала], unit as [ед. изм.], countAtStore as [количество на складе] from materials where countAtStore = {numericUpDown1.Value}";
                }
            }
            DataTable dt = GetData(filter);
            dataGridView1.DataSource = dt;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.AutoResizeColumns();
            button6.Enabled = true;
        }
        private void button6_Click(object sender, EventArgs e) {
            materials();
            button6.Enabled = false;
        }
        private void Form11_Shown(object sender, EventArgs e) {
            numericUpDown1.DecimalPlaces = 2;
            numericUpDown1.Increment = 0.1M;
            sql = dataBase.getConnection();
            sql.Open();
            setMode();
            getAccess();
            materials();
        }
    }
}
