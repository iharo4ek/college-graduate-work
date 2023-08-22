using System;
using System.Globalization;
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
    public partial class Form16 : Form {
        private const string STRING_PATTERN = @"[^а-яА-яa-zA-Z]";
        private const string INT_PATTERN = @"[^0-9]";
        DataBase dataBase = DataBase.getInstance();
        User user = User.getInstance();
        Modes mode;
        private SqlConnection sql;
        private Form16() {
            InitializeComponent();
        }
        private static Form instance;
        public static Form getInstance() {
            if (instance == null) {
                instance = new Form16();
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
                numericUpDown1.Visible = false;
                numericUpDown3.Visible = false;
                numericUpDown2.Visible = false;
                comboBox3.Visible = false;
                label1.Visible = false;
                label2.Visible = false;
                label3.Visible = false;
                label4.Visible = false;
            }
        }
        private void plan() {
            string query = "select id_pos, productName as [товар], planMonth as [месяц], planYear as [год], countProducts as [количество, шт.] from productionPlan inner join products on productionPlan.id_product = products.id_product;";
            SqlDataAdapter adapter = new SqlDataAdapter(query, dataBase.getSrc());
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            dataGridView1.DataSource = dt;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.AutoResizeColumns();
        }
        void ComBx1() {
            string query = "select id_product, productName from products;";
            SqlDataAdapter adapter = new SqlDataAdapter(query, dataBase.getSrc());
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            comboBox3.DataSource = dt;
            comboBox3.ValueMember = "id_product";
            comboBox3.DisplayMember = "productName";
            comboBox1.DataSource = dt;
            comboBox1.ValueMember = "id_product";
            comboBox1.DisplayMember = "productName";
        }
        private void Form16_Load(object sender, EventArgs e) {
        }
        private void button2_Click(object sender, EventArgs e) {
            try {
                if (dataGridView1.Rows.Count == 0) { return; }
                string query = $"Delete From productionPlan Where id_pos = {dataGridView1.CurrentRow.Cells[0].Value}";
                SqlCommand sqlCommand = new SqlCommand(query, sql);
                sqlCommand.ExecuteNonQuery();
                MessageBox.Show("Данные успешно удалены", "SUCCESS", MessageBoxButtons.OK);
                plan();
            } catch (Exception ex) {
                MessageBox.Show($"Не удалось удалить данные", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
        private void Form16_FormClosing(object sender, FormClosingEventArgs e) {
            sql.Close();
            Form form = Form15.getInstance();
            form.Show();
        }
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e) {
            comboBox3.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            numericUpDown3.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            numericUpDown2.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            numericUpDown1.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
        }
        private void button3_Click(object sender, EventArgs e) {
            try {
                if (dataGridView1.Rows.Count == 0) { return; }
                if (numericUpDown1.Value <= 0) {
                    MessageBox.Show("Количество должно быть > 0", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                int month = DateTime.Today.Month;
                int year = DateTime.Today.Year;
                if (month > numericUpDown3.Value && year >= numericUpDown2.Value) {
                    MessageBox.Show("Вы не можете  сменить план за уже прошедший период", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                string s = numericUpDown1.Value.ToString().Replace(',','.');
                string query = $"Update productionPlan Set planMonth = {numericUpDown3.Value}, " +
                    $"planYear = {numericUpDown2.Value}, countProducts = {numericUpDown1.Value} " +
                    $"where id_pos = {dataGridView1.CurrentRow.Cells[0].Value}";
                SqlCommand com = new SqlCommand(query, sql);
                com.ExecuteNonQuery();
                plan();
                MessageBox.Show("Данные успешно изменены", "SUCCESS", MessageBoxButtons.OK);
            } catch (Exception ex) {
                MessageBox.Show($"Не удалось изменить данные Произошла ошибка", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void button1_Click(object sender, EventArgs e) {
            try {
                if (numericUpDown1.Value <= 0) {
                    MessageBox.Show("Количество должно быть > 0", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                SqlDataAdapter adapter2 = new SqlDataAdapter();
                DataTable table2 = new DataTable();
                string query2 = $"select * from productionPlan where id_product = {comboBox3.SelectedValue} and planMonth = {numericUpDown3.Value} and planYear = {numericUpDown2.Value}; ";
                SqlCommand command2 = new SqlCommand(query2, dataBase.getConnection());
                adapter2.SelectCommand = command2;
                adapter2.Fill(table2);
                if (table2.Rows.Count != 0) {
                    MessageBox.Show($"План производства для {comboBox3.Text} на {numericUpDown3.Value} месяц {numericUpDown2.Value} года уже существует", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                string query = $"Insert into productionPlan values ({comboBox3.SelectedValue},{numericUpDown3.Value},{numericUpDown2.Value}, {numericUpDown1.Value});";
                SqlCommand sqlCommand = new SqlCommand(query, sql);
                sqlCommand.ExecuteNonQuery();
                MessageBox.Show("Данные успешно добавлены", "SUCCESS", MessageBoxButtons.OK);
                plan();
            } catch (Exception ex) {
                MessageBox.Show($"{ex.Message}", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void checkBox2_CheckedChanged(object sender, EventArgs e) {
            if (checkBox2.Checked == true) {
                numericUpDown4.Enabled = true;
            } else {
                numericUpDown4.Enabled = false;
            }
            if (checkBox1.Checked == false && checkBox2.Checked == false) {
                plan();
            }
        }
        private void checkBox1_CheckedChanged(object sender, EventArgs e) {
            if (checkBox1.Checked == true) {
                numericUpDown5.Enabled = true;
            } else {
                numericUpDown5.Enabled = false;
            }
            if (checkBox2.Checked == false && checkBox3.Checked == false) {
                plan();
            }
        }
        private void checkBox3_CheckedChanged(object sender, EventArgs e) {
            if (checkBox3.Checked == true) {
                comboBox1.Enabled = true;
            } else {
                comboBox1.Enabled = false;
            }
            if (checkBox1.Checked == false && checkBox2.Checked == false) {
                plan();
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
        private void comboBox1_TextChanged(object sender, EventArgs e) {
            if (checkBox3.Checked) {
                string query = "";
                if (checkBox1.Checked == true) {
                    if (checkBox2.Checked == true) {
                        query = "select id_pos, productName as [товар], planMonth as [месяц], planYear as [год], " +
                            $"countProducts as [количество] from " +
                            $"productionPlan inner join products on productionPlan.id_product = products.id_product " +
                            $"where planMonth = {numericUpDown4.Value} and planYear = {numericUpDown5.Value} and productionPlan.id_product = {comboBox1.SelectedValue};";
                    } else {
                        query = "select id_pos, productName as [товар], planMonth as [месяц], planYear as [год], " +
                            $"countProducts as [количество] from " +
                            $"productionPlan inner join products on productionPlan.id_product = products.id_product " +
                            $"where planMonth = {numericUpDown4.Value} and productionPlan.id_product = {comboBox1.SelectedValue};";
                    }
                } else {
                    if (checkBox2.Checked == true) {
                        query = "select id_pos, productName as [товар], planMonth as [месяц], planYear as [год], " +
                            $"countProducts as [количество] from " +
                            $"productionPlan inner join products on productionPlan.id_product = products.id_product " +
                            $"where planYear = {numericUpDown5.Value} and productionPlan.id_product = {comboBox1.SelectedValue};";
                    } else {
                        query = "select id_pos, productName as [товар], planMonth as [месяц], planYear as [год], " +
                            $"countProducts as [количество] from " +
                            $"productionPlan inner join products on productionPlan.id_product = products.id_product " +
                            $"where productionPlan.id_product = {comboBox1.SelectedValue};";
                    }
                }
                filter(query);
            }
        }
        private void maskedTextBox3_TextChanged(object sender, EventArgs e) {
           
        }
        private void maskedTextBox4_TextChanged(object sender, EventArgs e) {

        }
        private void Form16_Shown(object sender, EventArgs e) {
            sql = dataBase.getConnection();
            sql.Open();
            numericUpDown1.Maximum = 20000;
            numericUpDown1.Minimum = 1000;
            numericUpDown3.Minimum = 1;
            numericUpDown3.Maximum = 12;
            numericUpDown4.Minimum = 1;
            numericUpDown4.Maximum = 12;
            numericUpDown2.Minimum = DateTime.Today.Year;
            numericUpDown2.Maximum = DateTime.Today.AddYears(1).Year;
            numericUpDown5.Minimum = DateTime.Today.Year;
            numericUpDown5.Maximum = DateTime.Today.AddYears(1).Year;
            ComBx1();
            setMode();
            getAccess();
            plan();
        }
        private void numericUpDown3_ValueChanged(object sender, EventArgs e) {
        }

        private void numericUpDown4_ValueChanged(object sender, EventArgs e) {
            if (checkBox2.Checked == true) {
                string query = "";
                if (checkBox1.Checked == true) {
                    if (checkBox3.Checked == true) {
                        query = "select id_pos, productName as [товар], planMonth as [месяц], planYear as [год], " +
                            $"countProducts as [количество] from " +
                            $"productionPlan inner join products on productionPlan.id_product = products.id_product " +
                            $"where planMonth = {numericUpDown4.Value} and planYear = {numericUpDown5.Value} and productionPlan.id_product = {comboBox1.SelectedValue};";
                    } else {
                        query = "select id_pos, productName as [товар], planMonth as [месяц], planYear as [год], " +
                            $"countProducts as [количество] from " +
                            $"productionPlan inner join products on productionPlan.id_product = products.id_product " +
                            $"where planMonth = {numericUpDown4.Value} and planYear = {numericUpDown5.Value};";
                    }
                } else {
                    if (checkBox3.Checked == true) {
                        query = "select id_pos, productName as [товар], planMonth as [месяц], planYear as [год], " +
                            $"countProducts as [количество] from " +
                            $"productionPlan inner join products on productionPlan.id_product = products.id_product " +
                            $"where planMonth = {numericUpDown4.Value} and productionPlan.id_product = {comboBox1.SelectedValue};";
                    } else {
                        query = "select id_pos, productName as [товар], planMonth as [месяц], planYear as [год], " +
                            $"countProducts as [количество] from " +
                            $"productionPlan inner join products on productionPlan.id_product = products.id_product " +
                            $"where planMonth = {numericUpDown4.Value};";
                    }
                }
                filter(query);
            }
        }

        private void numericUpDown5_ValueChanged(object sender, EventArgs e) {
            if (checkBox1.Checked) {
                string query = "";
                if (checkBox2.Checked == true) {
                    if (checkBox3.Checked == true) {
                        query = "select id_pos, productName as [товар], planMonth as [месяц], planYear as [год], " +
                            $"countProducts as [количество] from " +
                            $"productionPlan inner join products on productionPlan.id_product = products.id_product " +
                            $"where planMonth = {numericUpDown4.Value} and planYear = {numericUpDown5.Value} and productionPlan.id_product = {comboBox1.SelectedValue};";
                    } else {
                        query = "select id_pos, productName as [товар], planMonth as [месяц], planYear as [год], " +
                             $"countProducts as [количество] from " +
                             $"productionPlan inner join products on productionPlan.id_product = products.id_product " +
                             $"where planMonth = {numericUpDown4.Value} and planYear = {numericUpDown5.Value};";
                    }
                } else {
                    if (checkBox3.Checked == true) {
                        query = "select id_pos, productName as [товар], planMonth as [месяц], planYear as [год], " +
                            $"countProducts as [количество] from " +
                            $"productionPlan inner join products on productionPlan.id_product = products.id_product " +
                            $"where planYear = {numericUpDown4.Value} and productionPlan.id_product = {comboBox1.SelectedValue};";
                    } else {
                        query = "select id_pos, productName as [товар], planMonth as [месяц], planYear as [год], " +
                            $"countProducts as [количество] from " +
                            $"productionPlan inner join products on productionPlan.id_product = products.id_product " +
                            $"where planYear = {numericUpDown5.Value};";
                    }
                    filter(query);
                }
            }
        }
    }
}