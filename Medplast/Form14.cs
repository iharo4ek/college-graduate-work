using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
namespace Medplast {
    public partial class Form14 : Form {
        private Modes mode;
        private const string STRING_PATTERN = @"[^а-яА-яa-zA-Z ]";
        private const string INT_PATTERN = @"[^0-9,.]";
        DataBase dataBase = DataBase.getInstance();
        User user = User.getInstance();
        private SqlConnection sql;
        private string filter = $"";
        private Form14() {
            InitializeComponent();
        }
        private static Form instance;
        public static Form getInstance() {
            if (instance == null) {
                instance = new Form14();
            }
            return instance;
        }
        private void products() {
            string query = "select id_product, productName as [название продукции], productTypeName as[тип продукции], cost as [цена], countAtStore as [количество на складе]  from products inner join productTypes on products.id_productType = productTypes.id_productType;";
            SqlDataAdapter adapter = new SqlDataAdapter(query, dataBase.getSrc());
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            dataGridView1.DataSource = dt;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.AutoResizeColumns();
        }
        private void recipe() {
            string query = $"select id_pos, productName as [продукция], materialName as [материал], materialsCount as [количество материала] from recipe inner join products on recipe.id_product = products.id_product inner join materials on recipe.id_material = materials.id_material where products.id_product = {dataGridView1.CurrentRow.Cells[0].Value}";
            SqlDataAdapter adapter = new SqlDataAdapter(query, dataBase.getSrc());
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            dataGridView2.DataSource = dt;
            dataGridView2.Columns[0].Visible = false;
            dataGridView2.Columns[1].Visible = false;
            dataGridView2.AutoResizeColumns();
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
                case "мастер цеха": {
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
                numericUpDown1.Visible = false;
                comboBox2.Visible = false;
                label4.Visible = false;
                label1.Visible = false;
                comboBox4.Visible = false;
                numericUpDown3.Visible = false;
                label5.Visible = false;
                label6.Visible = false;
                button7.Visible = false;
                button8.Visible = false;
                button9.Visible = false;
                label2.Visible = false;
            }
        }
        void ComBx1() {
            string query = "select id_productType, productTypeName from productTypes;";
            SqlDataAdapter adapter = new SqlDataAdapter(query, dataBase.getSrc());
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            comboBox2.DataSource = dt;
            comboBox2.ValueMember = "id_productType";
            comboBox2.DisplayMember = "productTypeName";
            comboBox1.DataSource = dt;
            comboBox1.ValueMember = "id_productType";
            comboBox1.DisplayMember = "productTypeName";
        }
        void CombBx3() {
            string query = "select id_material, materialName from materials;";
            SqlDataAdapter adapter = new SqlDataAdapter(query, dataBase.getSrc());
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            comboBox4.DataSource = dt;
            comboBox4.ValueMember = "id_material";
            comboBox4.DisplayMember = "materialName";
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
        private void button2_Click(object sender, EventArgs e) {
            try {
                if (dataGridView1.Rows.Count == 0) { return; }
                string query = $"Delete From products Where id_product = {dataGridView1.CurrentRow.Cells[0].Value}";
                SqlCommand sqlCommand = new SqlCommand(query, sql);
                sqlCommand.ExecuteNonQuery();
                MessageBox.Show("Данные успешно удалены", "SUCCESS", MessageBoxButtons.OK);
                products();
                if (dataGridView1.Rows.Count == 0) {
                    dataGridView2.DataSource = null;
                }
                recipe();
            } catch (Exception ex) {
                MessageBox.Show($"Не удалось удалить данные: Произошла ошибка", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void button1_Click(object sender, EventArgs e) {
            try {
                if (textBox1.Text.Length < 6) {
                    MessageBox.Show("Слишком короткое название продукции", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (numericUpDown1.Value <= 0) {
                    MessageBox.Show("Цена за единицу должна быть больше 0", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                string s = numericUpDown1.Value.ToString();
                s = s.Replace(',','.');
                string query = $"Insert into products values (N'{textBox1.Text}', {comboBox2.SelectedValue} , " + s + ", 0);";
                SqlCommand sqlCommand = new SqlCommand(query, sql);
                sqlCommand.ExecuteNonQuery();
                MessageBox.Show("Данные успешно добавлены", "SUCCESS", MessageBoxButtons.OK);
                products();
            } catch (Exception ex) {
                MessageBox.Show($"{ex.Message}", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void Form14_FormClosing(object sender, FormClosingEventArgs e) {
            sql.Close();
            Form form = Form2.getInstance();
            form.Show();
        }
        private void button3_Click(object sender, EventArgs e) {
            try {
                if (dataGridView1.Rows.Count == 0) { return; }
                if (textBox1.Text.Length < 6) {
                    MessageBox.Show("Слишком короткое название продукции", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (numericUpDown1.Value <= 0) {
                    MessageBox.Show("Цена за единицу должна быть больше 0", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                string query = $"Update products Set productName = N'{textBox1.Text}', id_productType = {comboBox2.SelectedValue}, " +
                    $"cost = {numericUpDown1.Value}  where id_product = {dataGridView1.CurrentRow.Cells[0].Value}";
                SqlCommand com = new SqlCommand(query, sql);
                com.ExecuteNonQuery();
                products();
                MessageBox.Show("Данные успешно изменены", "SUCCESS", MessageBoxButtons.OK);
            } catch (Exception ex) {
                MessageBox.Show($"Не удалось изменить данные: Произошла ошибка", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void экспортВExcelToolStripMenuItem_Click(object sender, EventArgs e) {
            try {
                Excel.Application exelApp = new Excel.Application();
                exelApp.Workbooks.Add();
                Excel.Worksheet wsh = (Excel.Worksheet)exelApp.ActiveSheet;
                wsh.Rows[1].Style.Font.Size = 12;
                exelApp.Cells[1, 1] = "продукция производимая ОАО Медпласт";
                wsh.Range[wsh.Cells[1, 1], wsh.Cells[2, dataGridView1.Rows[0].Cells.Count - 1]].Merge();
                for (int i = 0; i < dataGridView1.RowCount; i++) {
                    for (int j = 1; j < dataGridView1.ColumnCount; j++) {
                        wsh.Columns.AutoFit();
                        if (j == 3) {
                            wsh.Cells[3, j] = dataGridView1.Columns[j].HeaderText.ToString();
                            wsh.Cells[i + 4, j] = double.Parse(dataGridView1[j, i].Value.ToString());
                        } else {
                            wsh.Cells[3, j] = dataGridView1.Columns[j].HeaderText.ToString();
                            wsh.Cells[i + 4, j] = dataGridView1[j, i].Value.ToString();
                        }
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
        private void сформироватьОстаткиПоСкладуToolStripMenuItem_Click(object sender, EventArgs e) {
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
                DataTable dt = GetData($"select productName, cost, countAtStore from products");
                Word.Table tb = wordDocument.Tables[1];
                ReplaceWordStub("{date}", DateTime.Now.ToString("dd.MM.yyyy"), wordDocument);
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
                ReplaceWordStub("{date}", DateTime.Now.ToString(), wordDocument);
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
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e) {
            textBox1.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            numericUpDown1.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            comboBox2.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            comboBox4.Enabled = true;
            numericUpDown3.Enabled = true;
            button7.Enabled = true;
            button8.Enabled = true;
            button9.Enabled = true;
            recipe();
        }
        private void button6_Click(object sender, EventArgs e) {
            products();
            button6.Enabled = false;
        }
        private void button4_Click(object sender, EventArgs e) {
            if (radioButton1.Checked == true) {
                if (textBox3.Text.Length <= 0) {
                    MessageBox.Show($"Для фильтрации необходимо заполнить данные", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                } else {
                    filter = $"select id_product, productName as [название продукции], productTypeName as[тип продукции], cost as [цена], countAtStore as [количество на складе]  from products inner join productTypes on products.id_productType = productTypes.id_productType where productName = N'{textBox3.Text}'";
                }
            }
            if (radioButton2.Checked == true) {
                filter = $"select id_product, productName as [название продукции], productTypeName as[тип продукции], cost as [цена], countAtStore as [количество на складе]  from products inner join productTypes on products.id_productType = productTypes.id_productType where products.id_productType = {comboBox1.SelectedValue};";
            }
            if (radioButton3.Checked == true) {
                if (numericUpDown2.Value < 0) {
                    MessageBox.Show($"Цена не может быть меньше 0", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                } else {
                    string s = numericUpDown2.Value.ToString();
                    s = s.Replace(',', '.');
                    filter = $"select id_product, productName as [название продукции], productTypeName as[тип продукции], cost as [цена], countAtStore as [количество на складе]  from products inner join productTypes on products.id_productType = productTypes.id_productType where cost = " + s + ";";
                }
            }
            DataTable dt = GetData(filter);
            dataGridView1.DataSource = dt;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.AutoResizeColumns();
            button6.Enabled = true;
        }
        private void radioButton1_CheckedChanged(object sender, EventArgs e) {
            if (radioButton1.Checked == true) {
                button4.Enabled = true;
                textBox3.Text = "";
                textBox3.Visible = true;
            }
        }
        private void radioButton2_CheckedChanged(object sender, EventArgs e) {
            if (radioButton2.Checked == true) {
                button4.Enabled = true;
                comboBox1.Visible = true;
            }
        }
        private void radioButton3_CheckedChanged(object sender, EventArgs e) {
            if (radioButton3.Checked == true) {
                button4.Enabled = true;
                numericUpDown2.Value = 0;
                numericUpDown2.Visible = true;
            }
        }
        private void button7_Click(object sender, EventArgs e) {
            try {
                DataTable dataTable = GetData($"select * from recipe where id_product = {dataGridView1.CurrentRow.Cells[0].Value} and id_material = {comboBox4.SelectedValue};");
                if (dataTable.Rows.Count != 0) {
                    MessageBox.Show("Данный материал уже есть в рецепте", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (numericUpDown3.Value <= 0) {
                    MessageBox.Show("Количество должно быть больше 0", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                string s = numericUpDown3.Value.ToString();
                s = s.Replace(',', '.');
                string query = $"insert into recipe (id_product, id_material, materialsCount) values ({dataGridView1.CurrentRow.Cells[0].Value}, {comboBox4.SelectedValue}, " + s + ")";
                SqlCommand sqlCommand = new SqlCommand(query, sql);
                sqlCommand.ExecuteNonQuery();
                MessageBox.Show("Данные успешно добавлены", "SUCCESS", MessageBoxButtons.OK);
                recipe();
            } catch (Exception ex) {
                MessageBox.Show($"{ex.Message}", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void button8_Click(object sender, EventArgs e) {
            try {
                if (dataGridView2.Rows.Count == 0) { return; }
                string query = $"Delete From recipe Where id_pos = {dataGridView2.CurrentRow.Cells[0].Value}";
                SqlCommand sqlCommand = new SqlCommand(query, sql);
                sqlCommand.ExecuteNonQuery();
                MessageBox.Show("Данные успешно удалены", "SUCCESS", MessageBoxButtons.OK);
                recipe();
            } catch (Exception ex) {
                MessageBox.Show($"Не удалось удалить данные", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void button9_Click(object sender, EventArgs e) {
            try {
                if (dataGridView2.Rows.Count == 0) { return; }
                if (numericUpDown3.Value <= 0) {
                    MessageBox.Show("Количество должно быть больше 0", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                string s = numericUpDown3.Value.ToString();
                s = s.Replace(',','.');
                string query = $"Update recipe Set id_product = {dataGridView1.CurrentRow.Cells[0].Value}, id_material = {comboBox4.SelectedValue}, materialsCount = " + s + "  where id_pos = {dataGridView2.CurrentRow.Cells[0].Value}";
                SqlCommand com = new SqlCommand(query, sql);
                com.ExecuteNonQuery();
                recipe();
                MessageBox.Show("Данные успешно изменены", "SUCCESS", MessageBoxButtons.OK);
            } catch (Exception ex) {
                MessageBox.Show($"Не удалось изменить данные Произошла ошибка", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void Form14_Shown(object sender, EventArgs e) {
            //CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("en-US");
            numericUpDown1.DecimalPlaces = 2;
            numericUpDown1.Increment = 0.1M;
            numericUpDown2.DecimalPlaces = 2;
            numericUpDown2.Increment = 0.1M;
            numericUpDown3.DecimalPlaces = 2;
            numericUpDown3.Increment = 0.1M;
            sql = dataBase.getConnection();
            sql.Open();
            setMode();
            getAccess();
            products();
            ComBx1();
            CombBx3();
        }
        private void Form14_Load(object sender, EventArgs e) {
        }
    }
}
