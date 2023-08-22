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
    public partial class Form15 : Form {
        DataBase dataBase = DataBase.getInstance();
        private Modes mode;
        private const string STRING_PATTERN = @"[^а-яА-яa-zA-Z]";
        private const string INT_PATTERN = @"[^0-9]";
        private SqlConnection sql;
        private bool hiding = false;
        User user = User.getInstance();
        private DataTable dataTable;
        private Form15() {
            InitializeComponent();
        }
        private static Form instance;
        public static Form getInstance() {
            if (instance == null) {
                instance = new Form15();
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
                case "бухгалтер": {
                        this.mode = Modes.READ;
                        break;
                    }
                case "главный бухгалтер": {
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
                button4.Visible = false;
                numericUpDown1.Visible = false;
                comboBox4.Visible = false;
                label4.Visible = false;
                label1.Visible = false;
            }
        }
        void ComBx4() {
            string query = "select id_product, productName from products";
            SqlDataAdapter adapter = new SqlDataAdapter(query, dataBase.getSrc());
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            comboBox4.DataSource = dt;
            comboBox4.ValueMember = "id_product";
            comboBox4.DisplayMember = "productName";
            comboBox1.DataSource = dt;
            comboBox1.ValueMember = "id_product";
            comboBox1.DisplayMember = "productName";
        }
        private void production() {
            string query = "select id_pos, dateProduction as [дата], productName as [продукция], countProducts as [количество, шт.] from production inner join products on production.id_product = products.id_product;";
            SqlDataAdapter adapter = new SqlDataAdapter(query, dataBase.getSrc());
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            dataGridView1.DataSource = dt;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.AutoResizeColumns();
        }
        private void Form15_Load(object sender, EventArgs e) {
        }
        private void Form15_FormClosing(object sender, FormClosingEventArgs e) {
            sql.Close();
        }
        private void button6_Click(object sender, EventArgs e) {
            for (int i = 0; i < dataGridView1.RowCount; i++) {
                dataGridView1.Rows[i].Selected = false;
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    if (dataGridView1.Rows[i].Cells[j].Value != null)
                        if (dataGridView1.Rows[i].Cells[j].Value.ToString().ToLower().Contains(textBox1.Text.ToLower())) {
                            dataGridView1.Rows[i].Selected = true;
                            break;
                        }
            }
        }
        private void button4_Click(object sender, EventArgs e) {
            try {
                dataTable = GetData($"select * from productionPlan where id_product = {comboBox4.SelectedValue}");
                if (dataTable.Rows.Count == 0) {
                    MessageBox.Show("Данный товар отсутствует в плане производства", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                dataTable = GetData($"select * from production where id_product = {comboBox4.SelectedValue} and dateProduction = '{dateTimePicker1.Value.ToString("yyyy-MM-dd")}'");
                if (dataTable.Rows.Count != 0) {
                    MessageBox.Show("Данный товар уже учтен в производстве за сегодня", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (numericUpDown1.Value <= 0) {
                    MessageBox.Show("Количество произведенной продукции должно быть больше 0", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                dataTable = GetData($"select materialsCount, countAtStore from recipe inner join materials on recipe.id_material = materials.id_material where id_product = {comboBox4.SelectedValue}");
                if (dataTable.Rows.Count == 0) {
                    MessageBox.Show("Для данной продукции отсутствует рецепт", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                for (int i = 0; i < dataTable.Rows.Count; i++) {
                    int count2 = int.Parse(numericUpDown1.Value.ToString());
                    double need = count2 * double.Parse(dataTable.Rows[i].ItemArray[0].ToString());
                    double atStore = double.Parse(dataTable.Rows[i].ItemArray[1].ToString());
                    if (need > atStore) {
                        MessageBox.Show("Для производства не хватает материалов", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                dataTable = GetData($"select countProducts from productionPlan where id_product = {comboBox4.SelectedValue} and planMonth = {DateTime.Today.Month} and planYear = {DateTime.Today.Year}");
                int count = int.Parse(dataTable.Rows[0].ItemArray[0].ToString());
                count /= 30;
                if (numericUpDown1.Value == count || numericUpDown1.Value + 100 >= count || numericUpDown1.Value + 100 <= count || numericUpDown1.Value - 100 >= count || numericUpDown1.Value - 100 <= count) {
                    string query = $"Insert into production values ('{dateTimePicker1.Value.ToString("yyyy-MM-dd")}', {comboBox4.SelectedValue}, {numericUpDown1.Value});";
                    SqlCommand sqlCommand = new SqlCommand(query, sql);
                    if (sql.State == ConnectionState.Closed) {
                        sql.Open();
                    }
                    sqlCommand.ExecuteNonQuery();
                    MessageBox.Show("Данные успешно добавлены", "SUCCESS", MessageBoxButtons.OK);
                    production();
                }
            } catch (Exception ex) {
                MessageBox.Show($"{ex.Message}", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void filter(string query) {
            SqlDataAdapter adapter = new SqlDataAdapter(query, dataBase.getSrc());
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            dataGridView1.DataSource = dt;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.AutoResizeColumns();
        }
        private void dateTimePicker2_ValueChanged(object sender, EventArgs e) {
            if (checkBox2.Checked == true) {
                string query = "";
                if (checkBox1.Checked) {
                    if (checkBox3.Checked) {
                        query = $"select id_pos, dateProduction as [дата], productName as [продукция], " +
                            $"countProducts as [количество] from production inner join products on " +
                            $"production.id_product = products.id_product " +
                            $"where dateProduction >= '{dateTimePicker2.Value.ToString("yyyy - MM - dd")}' and dateProduction <= " +
                            $"'{dateTimePicker3.Value.ToString("yyyy - MM - dd")}' and production.id_product = {comboBox1.SelectedValue}";
                    } else {
                        query = $"select id_pos, dateProduction as [дата], productName as [продукция], " +
                            $"countProducts as [количество] from production inner join products on " +
                            $"production.id_product = products.id_product " +
                            $"where dateProduction >= '{dateTimePicker2.Value.ToString("yyyy - MM - dd")}' and dateProduction <= " +
                            $"'{dateTimePicker3.Value.ToString("yyyy - MM - dd")}'";
                    }
                } else {
                    if (checkBox3.Checked) {
                        query = $"select id_pos, dateProduction as [дата], productName as [продукция], " +
                            $"countProducts as [количество] from production inner join products on " +
                            $"production.id_product = products.id_product " +
                            $"where dateProduction >= '{dateTimePicker2.Value.ToString("yyyy - MM - dd")}' " +
                            $"and production.id_product = {comboBox1.SelectedValue}";
                    } else {
                        query = $"select id_pos, dateProduction as [дата], productName as [продукция], " +
                            $"countProducts as [количество] from production inner join products on " +
                            $"production.id_product = products.id_product " +
                            $"where dateProduction >= '{dateTimePicker2.Value.ToString("yyyy - MM - dd")}';";
                    }
                }
                filter(query);
            }
        }
        private void checkBox2_CheckedChanged(object sender, EventArgs e) {
            if (checkBox2.Checked == true) {
                dateTimePicker2.Enabled = true;
            } else {
                dateTimePicker2.Enabled = false;
            }
            if (checkBox1.Checked == false && checkBox3.Checked == false) {
                production();
            }
        }
        private void checkBox1_CheckedChanged(object sender, EventArgs e) {
            if (checkBox1.Checked == true) {
                dateTimePicker3.Enabled = true;
            } else {
                dateTimePicker3.Enabled = false;
            }
            if (checkBox2.Checked == false && checkBox3.Checked == false) {
                production();
            }
        }
        private void comboBox1_TextChanged(object sender, EventArgs e) {
            if (checkBox3.Checked == true) {
                string query = "";
                if (checkBox2.Checked == true) {
                    if (checkBox1.Checked == true) {
                        query = "select id_pos, dateProduction as [дата], productName as [продукция], countProducts as [количество] " +
                            $"from production inner join products on production.id_product = products.id_product " +
                            $"where dateProduction >= '{dateTimePicker2.Value.ToString("yyyy - MM - dd")}' and dateProduction <= '{dateTimePicker3.Value.ToString("yyyy - MM - dd")}' " +
                            $"and production.id_product = {comboBox1.SelectedValue};";
                    } else {
                        query = "select id_pos, dateProduction as [дата], productName as [продукция], countProducts as [количество] " +
                            $"from production inner join products on production.id_product = products.id_product " +
                            $"where dateProduction >= '{dateTimePicker2.Value.ToString("yyyy - MM - dd")}' " +
                            $"and production.id_product = {comboBox1.SelectedValue};";
                    }
                } else {
                    if (checkBox1.Checked == true) {
                        query = "select id_pos, dateProduction as [дата], productName as [продукция], countProducts as [количество] " +
                            $"from production inner join products on production.id_product = products.id_product " +
                            $"where dateProduction <= '{dateTimePicker3.Value.ToString("yyyy - MM - dd")}' " +
                            $"and production.id_product = {comboBox1.SelectedValue};";
                    } else {
                        query = "select id_pos, dateProduction as [дата], productName as [продукция], countProducts as [количество] " +
                            $"from production inner join products on production.id_product = products.id_product " +
                            $"where production.id_product = {comboBox1.SelectedValue};";
                    }
                }
                filter(query);
            }
        }
        private void textBox1_TextChanged(object sender, EventArgs e) {
        }
        private void dateTimePicker3_ValueChanged(object sender, EventArgs e) {
            if (checkBox1.Checked == true) {
                string query = "";
                if (checkBox2.Checked == true) {
                    if (checkBox3.Checked == true) {
                        query = "select id_pos, dateProduction as [дата], productName as [продукция], countProducts as [количество] " +
                            $"from production inner join products on production.id_product = products.id_product " +
                            $"where dateProduction >= '{dateTimePicker2.Value.ToString("yyyy - MM - dd")}' dateProduction and <= '{dateTimePicker3.Value.ToString("yyyy - MM - dd")}'" +
                            $" and production.id_product = {comboBox1.SelectedValue};";
                    } else {
                        query = "select id_pos, dateProduction as [дата], productName as [продукция], countProducts as [количество] " +
                            $"from production inner join products on production.id_product = products.id_product " +
                            $"where dateProduction >= '{dateTimePicker2.Value.ToString("yyyy - MM - dd")}' and dateProduction <= '{dateTimePicker3.Value.ToString("yyyy - MM - dd")}'";
                    }
                } else {
                    if (checkBox3.Checked == true) {
                        query = "select id_pos, dateProduction as [дата], productName as [продукция], countProducts as [количество] " +
                            $"from production inner join products on production.id_product = products.id_product " +
                            $"where dateProduction <= '{dateTimePicker3.Value.ToString("yyyy - MM - dd")}'" +
                            $" and production.id_product = {comboBox1.SelectedValue};";
                    } else {
                        query = "select id_pos, dateProduction as [дата], productName as [продукция], countProducts as [количество] " +
                            $"from production inner join products on production.id_product = products.id_product " +
                            $"where dateProduction <= '{dateTimePicker3.Value.ToString("yyyy - MM - dd")}';";
                    }
                }
                filter(query);
            }
        }
        private void checkBox3_CheckedChanged(object sender, EventArgs e) {
            if (checkBox3.Checked == true) {
                comboBox1.Enabled = true;
            } else {
                comboBox1.Enabled = false;
            }
            if (checkBox1.Checked == false && checkBox2.Checked == false) {
                production();
            }
        }
        private void планПроизводстваToolStripMenuItem_Click(object sender, EventArgs e) {
            sql.Close();
            Form form = Form16.getInstance();
            this.Hide();
            form.ShowDialog();
        }
        private void Form15_FormClosed(object sender, FormClosedEventArgs e) {
            Form form = Form2.getInstance();
            form.Show();
        }
        private void ReplaceWordStub(string stubToReplace, string text, Word.Document wordDocumet) {
            var range = wordDocumet.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: stubToReplace, ReplaceWith: text);
        }
        public DataTable GetData(string query) {
            SqlDataAdapter adapter = new SqlDataAdapter(query, dataBase.getSrc());
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            return dt;
        }
        private void сформироватьОтчетОПроизводствеToolStripMenuItem_Click(object sender, EventArgs e) {
            if (dataGridView1.Rows.Count == 0) {
                MessageBox.Show($"Сначала нужно заполнить данные о  производтве", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (checkBox2.Checked == false) {
                MessageBox.Show($"Сначала нужно выбрать дату для формирования отчета", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            var wordapp = new Word.Application();
            string path = Environment.CurrentDirectory + @"\Otchet2.docx";
            var wordDocument = wordapp.Documents.Open(path);
            try {
                wordapp.Visible = false;
                DataTable dt = GetData($"select dateProduction, productName, countProducts, cost from production inner join products on production.id_product = products.id_product where dateProduction = '{dateTimePicker2.Value.ToString("yyyy - MM - dd")}';");
                DateTime dateTime = DateTime.Parse(dt.Rows[0].ItemArray[0].ToString());
                ReplaceWordStub("{date}", dateTime.ToString("dd.MM.yyyy"), wordDocument);
                int ind = 1, count = 0; double sum = 0;
                Word.Table tb = wordDocument.Tables[1];
                Word.Row r = tb.Rows[2];
                for (int index = 0; index < dt.Rows.Count; index++, ind++) {
                    r.Cells[1].Range.Text = ind.ToString();
                    r.Cells[2].Range.Text = dt.Rows[index].ItemArray[1].ToString();
                    r.Cells[3].Range.Text = dt.Rows[index].ItemArray[2].ToString();
                    r.Cells[4].Range.Text = dt.Rows[index].ItemArray[3].ToString();
                    double s = int.Parse(dt.Rows[index].ItemArray[2].ToString()) * double.Parse(dt.Rows[index].ItemArray[3].ToString());
                    r.Cells[5].Range.Text = s.ToString();
                    sum += s;
                    count += int.Parse(dt.Rows[index].ItemArray[2].ToString());
                    r = tb.Rows.Add();
                }
                wordDocument.Range(r.Cells[1].Range.Start, r.Cells[2].Range.End).Cells.Merge();
                r.Cells[1].Range.Text = "ИТОГО";
                r.Cells[4].Range.Text = sum.ToString();
                r.Cells[2].Range.Text = count.ToString();
                string fio = user.getSName() + " " + user.getName() + " " + user.getP();
                ReplaceWordStub("{empl}", fio, wordDocument);
                wordapp.Visible = true;
            } catch (Exception ex) {
                MessageBox.Show($"Произошла ошибка", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (wordDocument == null) { return; }
                wordDocument.Close();
            }
        }
        private void сформироватьактВнутреннегоПеремещенияToolStripMenuItem_Click(object sender, EventArgs e) {
            if (dataGridView1.Rows.Count == 0) {
                MessageBox.Show($"Сначала нужно заполнить данные о  производтве", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (checkBox2.Checked == false) {
                MessageBox.Show($"Сначала нужно выбрать дату для формирования отчета", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            var wordapp = new Word.Application();
            string path = Environment.CurrentDirectory + @"\Akt.docx";
            var wordDocument = wordapp.Documents.Open(path);
            try {
                wordapp.Visible = false;
                DataTable dt = GetData($"select dateProduction, materialName, unit, round(sum(materialsCount*countProducts),2) from production inner join products on production.id_product = products.id_product inner join recipe on products.id_product = recipe.id_product inner join materials on recipe.id_material = materials.id_material where dateProduction = '{dateTimePicker2.Value.ToString("yyyy - MM - dd")}' group by dateProduction,materialName,unit; ");
                DateTime dateTime = DateTime.Parse(dt.Rows[0].ItemArray[0].ToString());
                ReplaceWordStub("{date}", dateTime.ToString("dd.MM.yyyy"), wordDocument);
                Word.Table tb = wordDocument.Tables[3];
                Word.Row r = tb.Rows[4];
                double count = 0;
                for (int index = 0; index < dt.Rows.Count; index++) {
                    r.Cells[3].Range.Text = dt.Rows[index].ItemArray[1].ToString();
                    r.Cells[6].Range.Text = dt.Rows[index].ItemArray[2].ToString();
                    r.Cells[7].Range.Text = dt.Rows[index].ItemArray[3].ToString();
                    r.Cells[8].Range.Text = dt.Rows[index].ItemArray[3].ToString();
                    count += double.Parse(dt.Rows[index].ItemArray[3].ToString());
                    r = tb.Rows.Add();
                }
                r.Cells[1].Range.Text = "ИТОГО";
                r.Cells[7].Range.Text = count.ToString();
                r.Cells[8].Range.Text = count.ToString();
                wordapp.Visible = true;
            } catch (Exception ex) {
                MessageBox.Show($"Произошла ошибка", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (wordDocument == null) { return; }
                wordDocument.Close();
            }
        }
        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e) {
            dataTable = GetData($"select countProducts from productionPlan where id_product = {comboBox4.SelectedValue} and planMonth = {DateTime.Today.Month} and planYear = {DateTime.Today.Year}");
            if (dataTable.Rows.Count == 0) {
                return;
            }
            int count = int.Parse(dataTable.Rows[0].ItemArray[0].ToString());
            count /= 30;
            numericUpDown1.Minimum = count - 100;
            numericUpDown1.Maximum = count + 100;
        }
        private void numericUpDown1_ValueChanged(object sender, EventArgs e) {
            if (numericUpDown1.Value > numericUpDown1.Maximum) {
                numericUpDown1.Value = numericUpDown1.Maximum;
            }
            if (numericUpDown1.Value < numericUpDown1.Minimum) {
                numericUpDown1.Value = numericUpDown1.Minimum;
            }
        }
        private void Form15_Shown(object sender, EventArgs e) {
            comboBox4.SelectedIndexChanged -= comboBox4_SelectedIndexChanged;
            sql = dataBase.getConnection();
            sql.Open();
            dateTimePicker1.MinDate = DateTime.Today;
            dateTimePicker1.MaxDate = DateTime.Today;
            setMode();
            getAccess();
            production();
            ComBx4();
            comboBox4.SelectedIndexChanged += comboBox4_SelectedIndexChanged;
            dataTable = GetData($"select countProducts from productionPlan where id_product = {comboBox4.SelectedValue} and planMonth = {DateTime.Today.Month} and planYear = {DateTime.Today.Year}");
            if (dataTable.Rows.Count == 0) {
                return;
            }
            int count = int.Parse(dataTable.Rows[0].ItemArray[0].ToString());
            count /= 30;
            numericUpDown1.Minimum = count - 100;
            numericUpDown1.Maximum = count + 100;
        }
        private void отчетОВыполненииПланаЗаМесяцToolStripMenuItem_Click(object sender, EventArgs e) {
            var wordapp = new Word.Application();
            string path = Environment.CurrentDirectory + @"\Otchet3.docx";
            var wordDocument = wordapp.Documents.Open(path);
            try {
                if (checkBox2.Checked == false) { return; }
                DataTable dt = GetData($"select productName, round(SUM(productionPlan.countProducts),2), round(SUM(productionPlan.countProducts*products.cost),2), round(Sum(production.countProducts),2), round(sum(production.countProducts*products.cost),2), round(sum(production.countProducts - productionPlan.countProducts),2),products.id_product from production inner join products on production.id_product = products.id_product inner join productionPlan on products.id_product = productionPlan.id_product where planMonth = {DateTime.Parse(dateTimePicker2.Value.ToString()).Month} and planYear = {DateTime.Parse(dateTimePicker2.Value.ToString()).Year} group by productName, products.id_product ");
                if (dt.Rows.Count == 0) { return; }
                wordapp.Visible = false;
                Word.Table tb = wordDocument.Tables[1];
                Word.Row r = tb.Rows[2];
                ReplaceWordStub("{month}", DateTime.Parse(dateTimePicker2.Value.ToString()).Month.ToString(), wordDocument);
                ReplaceWordStub("{y}", DateTime.Parse(dateTimePicker2.Value.ToString()).Year.ToString(), wordDocument);
                for (int i = 0; i < dt.Rows.Count; i++) {
                    r.Cells[1].Range.Text = dt.Rows[i].ItemArray[0].ToString();
                    r.Cells[2].Range.Text = dt.Rows[i].ItemArray[1].ToString();
                    r.Cells[3].Range.Text = dt.Rows[i].ItemArray[2].ToString();
                    r.Cells[4].Range.Text = dt.Rows[i].ItemArray[3].ToString();
                    r.Cells[5].Range.Text = dt.Rows[i].ItemArray[4].ToString();
                    double d = double.Parse(dt.Rows[i].ItemArray[5].ToString());
                    if (d < 0) {
                        r.Cells[6].Range.Text = "-";
                        r.Cells[7].Range.Text = dt.Rows[i].ItemArray[5].ToString();
                    } else if (d > 0) {
                        r.Cells[7].Range.Text = "-";
                        r.Cells[6].Range.Text = dt.Rows[i].ItemArray[5].ToString();
                    } else {
                        r.Cells[6].Range.Text = "-";
                        r.Cells[7].Range.Text = "-";
                    }
                    string s = $"{DateTime.Parse(dateTimePicker2.Value.ToString()).Year.ToString()}-{DateTime.Parse(dateTimePicker2.Value.ToString()).Month.ToString("")}-01";
                    string s2 = $"{DateTime.Parse(dateTimePicker2.Value.ToString()).Year.ToString()}-{DateTime.Parse(dateTimePicker2.Value.ToString()).Month.ToString("")}-31";
                    DateTime dateTime = DateTime.Parse(s);
                    string m = "";
                    if (dateTime.Month < 10) { m = "0" + dateTime.Month.ToString(); }
                    s = dateTime.Year + "-" + m + "-01";
                    s2 = dateTime.Year + "-" + m + "-"+DateTime.DaysInMonth(dateTime.Year,int.Parse(m));
                    DataTable dt2 = GetData($"select SUM(saleTemp.countProducts) from sale inner join saleTemp on sale.id_sale = saleTemp.id_sale inner join products on saleTemp.id_product = products.id_product where products.id_product = {dt.Rows[i].ItemArray[6].ToString()} and  saleDate >= '{s}' and saleDate <= '{s2}' group by productName");
                    if (dt2.Rows.Count == 0) {
                        r = tb.Rows.Add();
                        continue;
                    }
                    r.Cells[8].Range.Text = dt2.Rows[0].ItemArray[0].ToString();
                    r = tb.Rows.Add();
                }
                wordapp.Visible = true;
            } catch {
                MessageBox.Show($"Произошла ошибка", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (wordDocument == null) { return; }
                wordDocument.Close();
            }
        }
    }
}
