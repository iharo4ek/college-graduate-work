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
using Word = Microsoft.Office.Interop.Word;
namespace Medplast {
    public partial class Form9 : Form {
        private Modes mode;
        private const string STRING_PATTERN = @"[^а-яА-яa-zA-Z]";
        private const string INT_PATTERN = @"[^0-9]";
        DataBase dataBase = DataBase.getInstance();
        User user = User.getInstance();
        private static SqlConnection sql;
        private Form9() {
            InitializeComponent();
        }
        private static Form instance;
        public static Form getInstance() {
            if (instance == null) {
                instance = new Form9();
            }
            return instance;
        }
        private void sale() {
            string query = "select id_sale, (employeeSurname + ' ' + employeeName + ' ' + employeePatronymic) as [мэнеджер], saleDate as [дата продажи], nameOrganisation as [клиент], status as [статус продажи], summ as [стоимость, бел.руб.], numberOfTheCar [машина доставки], departureDate as [дата отправки] from sale left join clients on sale.id_client = clients.id_client left join cars on sale.id_car = cars.id_car left join employees on sale.id_manager = employees.id_employee;";
            SqlDataAdapter adapter = new SqlDataAdapter(query, dataBase.getSrc());
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            dataGridView1.DataSource = dt;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.AutoResizeColumns();
        }
        private void saleTemp() {
            if (dataGridView1.CurrentRow == null) {
                return;
            }
            string query = $"select id_pos, id_sale, productName as [продукция], countProducts as [количество продукции, шт] from saleTemp inner join products on saleTemp.id_product = products.id_product where id_sale = {dataGridView1.CurrentRow.Cells[0].Value};";
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
                button6.Visible = false;
                comboBox1.Visible = false;
                comboBox3.Visible = false;
                comboBox4.Visible = false;
                dateTimePicker2.Visible = false;
                numericUpDown1.Visible = false;
                label1.Visible = false;
                label3.Visible = false;
                label4.Visible = false;
                label5.Visible = false;
                label6.Visible = false;
            }
        }
        void ComBx1() {
            string query = "select id_client, nameOrganisation from clients;";
            SqlDataAdapter adapter = new SqlDataAdapter(query, dataBase.getSrc());
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            comboBox1.DataSource = dt;
            comboBox1.ValueMember = "id_client";
            comboBox1.DisplayMember = "nameOrganisation";
        }
        void ComBx3() {
            string query = "select id_car, numberOfTheCar from cars;";
            SqlDataAdapter adapter = new SqlDataAdapter(query, dataBase.getSrc());
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            comboBox3.DataSource = dt;
            comboBox3.ValueMember = "id_car";
            comboBox3.DisplayMember = "numberOfTheCar";
        }
        void ComBx4() {
            string query = "select id_product, (productName + ' ' + STR(countAtStore) + N'шт') as [pr] from products;";
            SqlDataAdapter adapter = new SqlDataAdapter(query, dataBase.getSrc());
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            comboBox4.DataSource = dt;
            comboBox4.ValueMember = "id_product";
            comboBox4.DisplayMember = "pr";
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
        private void button4_Click(object sender, EventArgs e) {
            for (int i = 0; i < dataGridView2.RowCount; i++) {
                dataGridView2.Rows[i].Selected = false;
                for (int j = 0; j < dataGridView2.ColumnCount; j++)
                    if (dataGridView2.Rows[i].Cells[j].Value != null)
                        if (dataGridView2.Rows[i].Cells[j].Value.ToString().ToLower().Contains(textBox3.Text.ToLower())) {
                            dataGridView2.Rows[i].Selected = true;
                            break;
                        }
            }
        }
        private void Form9_FormClosing(object sender, FormClosingEventArgs e) {
            sql.Close();
            Form form = Form2.getInstance();
            form.Show();
        }
        private void button2_Click(object sender, EventArgs e) {
            if (dataGridView1.Rows.Count == 0) {
                return;
            }
            try {
                string query = $"Delete From sale Where id_sale = {dataGridView1.CurrentRow.Cells[0].Value}";
                SqlCommand sqlCommand = new SqlCommand(query, sql);
                sqlCommand.ExecuteNonQuery();
                MessageBox.Show("Данные успешно удалены", "SUCCESS", MessageBoxButtons.OK);
                sale();
                saleTemp();
                if (dataGridView1.Rows.Count == 0) {
                    dataGridView2.DataSource = null;
                }
            } catch (Exception ex) {
                MessageBox.Show($"Не удалось удалить данные, попробуйте удалить данные из связанных таблиц", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void button1_Click(object sender, EventArgs e) {
            if (comboBox1.Enabled == false) {
                comboBox1.Enabled = true;
                comboBox3.Enabled = false;
                dateTimePicker2.Enabled = false;
            } else {
                try {
                    DataTable dataTable = GetData($"select * from sale where id_client = {comboBox1.SelectedValue} and saleDate = '{dateTimePicker1.Value.ToString("yyyy-MM-dd")}'");
                    if (dataTable.Rows.Count != 0) {
                        MessageBox.Show($"Заказ в это день уже оформлен на данного клиента", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    string query = $"Insert into sale values ({user.getId()}, '{dateTimePicker1.Value.ToString("yyyy-MM-dd")}', {comboBox1.SelectedValue}, N'оформлен' ,0,null, null);";
                    SqlCommand sqlCommand = new SqlCommand(query, sql);
                    sqlCommand.ExecuteNonQuery();
                    MessageBox.Show("Данные успешно добавлены", "SUCCESS", MessageBoxButtons.OK);
                    sale();
                    saleTemp();
                } catch (Exception ex) {
                    MessageBox.Show($"{ex.Message}", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e) {
            try {
                dateTimePicker2.Value = Convert.ToDateTime(dataGridView1.CurrentRow.Cells[7].Value.ToString());
                comboBox1.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                comboBox3.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
            } catch { }
            comboBox4.Enabled = true;
            numericUpDown1.Enabled = true;
            button6.Enabled = true;
            saleTemp();
        }
        private void button3_Click(object sender, EventArgs e) {
            if (dataGridView1.Rows.Count == 0) {
                return;
            }
            if (dataGridView1.CurrentRow == null) { return; }
            if (dataGridView1.CurrentRow.Cells[7].Value.ToString() != "") {
                return;
            }
            if (comboBox1.Enabled == true) {
                comboBox1.Enabled = false;
                comboBox3.Enabled = true;
                dateTimePicker2.Enabled = true;
                button1.Enabled = false;
            } else {
                try {
                    if (dataGridView1.Rows.Count == 0) { return; }
                    DataTable dataTable = GetData($"select * from materialsPurchase where id_car = {comboBox3.SelectedValue} and receivingDate = '{dateTimePicker2.Value.ToString("yyyy-MM-dd")}'");
                    if (dataTable.Rows.Count != 0) {
                        MessageBox.Show($"Данная машина занята в этот день", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    dataTable = GetData($"select * from sale where id_car = {comboBox3.SelectedValue} and departureDate = '{dateTimePicker2.Value.ToString("yyyy-MM-dd")}  '");
                    if (dataTable.Rows.Count != 0) {
                        MessageBox.Show($"Данная машина занята в этот день", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    if (dataGridView2.Rows.Count == 0) {
                        MessageBox.Show($"Сначало надо добавить данные в состав продажи", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    string query = $"Update sale Set id_car = {comboBox3.SelectedValue}, departureDate = '{dateTimePicker1.Value.ToString("yyyy-MM-dd")}', status = N'доставлен' where id_sale = {dataGridView1.CurrentRow.Cells[0].Value}";
                    SqlCommand com = new SqlCommand(query, sql);
                    com.ExecuteNonQuery();
                    sale();
                    MessageBox.Show("Данные успешно изменены", "SUCCESS", MessageBoxButtons.OK);
                    comboBox3.Enabled = false;
                    dataGridView2.Enabled = false;
                    button1.Enabled = true;
                    comboBox1.Enabled = true;
                } catch (Exception ex) {
                    MessageBox.Show($"Не удалось изменить данные Произошла ошибка", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        public DataTable GetData(string query) {
            SqlDataAdapter adapter = new SqlDataAdapter(query, dataBase.getSrc());
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            return dt;
        }
        private void button6_Click(object sender, EventArgs e) {
            if (dataGridView1.Rows.Count == 0) { return; }
            if (dataGridView1.CurrentRow.Cells[7].Value.ToString() != "") {
                MessageBox.Show("Вы не можете добавить товары в уже доставленный заказ", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (numericUpDown1.Value <= 0) {
                MessageBox.Show("Количество продукции должно быть больше 0", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            DataTable dataTable = GetData($"select id_product from saleTemp where id_product = {comboBox4.SelectedValue} and id_sale = {dataGridView1.CurrentRow.Cells[0].Value}");
            if (dataTable.Rows.Count != 0) {
                MessageBox.Show("Данный товар уже присутствует в продаже", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            try {
                DataTable dt2 = GetData($"select id_product, countAtStore from products where id_product = {comboBox4.SelectedValue};");
                if (int.Parse(dt2.Rows[0].ItemArray[1].ToString()) < numericUpDown1.Value) {
                    MessageBox.Show("На складе не хватает товаров", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                string query = $"Insert into saleTemp values ({dataGridView1.CurrentRow.Cells[0].Value}, {comboBox4.SelectedValue}, {numericUpDown1.Value});";
                SqlCommand sqlCommand = new SqlCommand(query, sql);
                sqlCommand.ExecuteNonQuery();
                MessageBox.Show("Данные успешно добавлены", "SUCCESS", MessageBoxButtons.OK);
                ComBx4();
                sale();
                saleTemp();
            } catch (Exception ex) {
                MessageBox.Show($"Произошла ошибка", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e) {
            comboBox4.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            numericUpDown1.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
        }
        private void ReplaceWordStub(string stubToReplace, string text, Word.Document wordDocumet) {
            var range = wordDocumet.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: stubToReplace, ReplaceWith: text);
        }
        private void сформироватьТТНToolStripMenuItem_Click(object sender, EventArgs e) {
            if (dataGridView1.Rows.Count == 0) {
                MessageBox.Show($"Сначала нужно добавить данные о продаже", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (dataGridView2.Rows.Count == 0) {
                MessageBox.Show($"Сначала нужно добавить данные о продаже", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (dataGridView1.CurrentRow.Cells[7].Value.ToString() == "") {
                MessageBox.Show("Для формирования ТТН необходимо доставить продажу","ERROR" ,MessageBoxButtons.OK ,MessageBoxIcon.Error);
                return;
            }
            var wordapp = new Word.Application();
            string path = Environment.CurrentDirectory + @"\SaleTTN.docx";
            var wordDocument = wordapp.Documents.Open(path);
            try {
                wordapp.Visible = false;
                DataTable dt = GetData($"select YNP, (carBrand + ' ' + numberOfTheCar) as [car], (employeeSurname + ' ' + employeeName + ' ' + employeePatronymic) as [driver], (nameOrganisation + ' ' + adress) as [client], adress as [adress],productName, cost, countProducts from sale inner join saleTemp on sale.id_sale = saleTemp.id_sale inner join clients on sale.id_client = clients.id_client inner join cars on sale.id_car = cars.id_car inner join products on saleTemp.id_product = products.id_product inner join employees on cars.id_driver = employees.id_employee where sale.id_sale = {dataGridView1.CurrentRow.Cells[0].Value.ToString()}");
                ReplaceWordStub("{clYNP}", dt.Rows[0].ItemArray[0].ToString(), wordDocument);
                ReplaceWordStub("{clYNP}", dt.Rows[0].ItemArray[0].ToString(), wordDocument);
                ReplaceWordStub("{Car}", dt.Rows[0].ItemArray[1].ToString(), wordDocument);
                ReplaceWordStub("{driver}", dt.Rows[0].ItemArray[2].ToString(), wordDocument);
                ReplaceWordStub("{client}", dt.Rows[0].ItemArray[3].ToString(), wordDocument);
                ReplaceWordStub("{client}", dt.Rows[0].ItemArray[3].ToString(), wordDocument);
                ReplaceWordStub("{id_sale}", dataGridView1.CurrentRow.Cells[0].Value.ToString(), wordDocument);
                ReplaceWordStub("{date}", DateTime.Today.ToString("dd.MM.yyyy"), wordDocument);
                ReplaceWordStub("{date}", DateTime.Today.ToString("dd.MM.yyyy"), wordDocument);
                ReplaceWordStub("{clA}", dt.Rows[0].ItemArray[4].ToString(), wordDocument);
                double stoim = 0; double sumNDS = 0; double stoimNDS = 0; int count = 0;
                Word.Table tb = wordDocument.Tables[2];
                Word.Row r = tb.Rows[4];
                for (int index = 0; index < dt.Rows.Count; index++) {
                    r.Cells[1].Range.Text = dt.Rows[index].ItemArray[5].ToString();
                    r.Cells[2].Range.Text = "шт";
                    r.Cells[3].Range.Text = dt.Rows[index].ItemArray[7].ToString();
                    r.Cells[4].Range.Text = dt.Rows[index].ItemArray[6].ToString();
                    r.Cells[5].Range.Text = (double.Parse(dt.Rows[index].ItemArray[6].ToString()) * int.Parse(dt.Rows[index].ItemArray[7].ToString())).ToString();
                    r.Cells[6].Range.Text = "20%";
                    r.Cells[7].Range.Text = ((double.Parse(dt.Rows[index].ItemArray[6].ToString()) * int.Parse(dt.Rows[index].ItemArray[7].ToString())) * 0.2).ToString();
                    r.Cells[8].Range.Text = (((double.Parse(dt.Rows[index].ItemArray[6].ToString()) * int.Parse(dt.Rows[index].ItemArray[7].ToString())) * 0.2) + double.Parse(dt.Rows[index].ItemArray[6].ToString()) * int.Parse(dt.Rows[index].ItemArray[7].ToString())).ToString();
                    r.Cells[9].Range.Text = dt.Rows[index].ItemArray[7].ToString();
                    count += int.Parse(dt.Rows[index].ItemArray[7].ToString());
                    stoim += double.Parse(dt.Rows[index].ItemArray[6].ToString()) * int.Parse(dt.Rows[index].ItemArray[7].ToString());
                    sumNDS += double.Parse(dt.Rows[index].ItemArray[6].ToString()) * int.Parse(dt.Rows[index].ItemArray[7].ToString()) * 0.2;
                    stoimNDS += ((double.Parse(dt.Rows[index].ItemArray[6].ToString()) * int.Parse(dt.Rows[index].ItemArray[7].ToString())) * 0.2) + double.Parse(dt.Rows[index].ItemArray[6].ToString()) * int.Parse(dt.Rows[index].ItemArray[7].ToString());
                    r = tb.Rows.Add();
                }
                r.Cells[1].Range.Text = "ИТОГО";
                r.Cells[2].Range.Text = "x";
                r.Cells[3].Range.Text = "{c}";
                r.Cells[4].Range.Text = "x";
                r.Cells[5].Range.Text = "{st}";
                r.Cells[6].Range.Text = "x";
                r.Cells[7].Range.Text = "{sum}";
                r.Cells[8].Range.Text = "{st2}";
                r.Cells[9].Range.Text = "{c}";
                r.Cells[10].Range.Text = "";
                r.Cells[11].Range.Text = "x";
                ReplaceWordStub("{c}", count.ToString(), wordDocument);
                ReplaceWordStub("{c}", count.ToString(), wordDocument);
                ReplaceWordStub("{st}", stoim.ToString(), wordDocument);
                ReplaceWordStub("{sum}", sumNDS.ToString(), wordDocument);
                ReplaceWordStub("{sum}", sumNDS.ToString(), wordDocument);
                ReplaceWordStub("{st2}", stoimNDS.ToString(), wordDocument);
                ReplaceWordStub("{st2}", stoimNDS.ToString(), wordDocument);
                wordapp.Visible = true;
            } catch (Exception ex) {
                MessageBox.Show($"Произошла ошибка", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (wordDocument == null) { return; }
                wordDocument.Close();
            }
        }
        private void cформироатьЗаявкуToolStripMenuItem_Click(object sender, EventArgs e) {
            if (dataGridView1.Rows.Count == 0) {
                MessageBox.Show($"Сначала нужно добавить данные о продаже", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (dataGridView2.Rows.Count == 0) {
                MessageBox.Show($"Сначала нужно добавить данные о продаже", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            var wordapp = new Word.Application();
            string path = Environment.CurrentDirectory + @"\Zayavka.docx";
            var wordDocument = wordapp.Documents.Open(path);
            try {
                wordapp.Visible = false;
                DataTable dt = GetData($"select saleDate, nameOrganisation, productName, countProducts from sale inner join saleTemp on sale.id_sale = saleTemp.id_sale inner join clients on sale.id_client = clients.id_client inner join products on saleTemp.id_product = products.id_product where sale.id_sale = {dataGridView1.CurrentRow.Cells[0].Value}");
                DateTime dateTime = DateTime.Parse(dt.Rows[0].ItemArray[0].ToString());
                ReplaceWordStub("{date}", dateTime.ToString("dd.MM.yyyy"), wordDocument);
                ReplaceWordStub("{client}", dt.Rows[0].ItemArray[1].ToString(), wordDocument);
                int ind = 1, count = 0;
                Word.Table tb = wordDocument.Tables[1];
                Word.Row r = tb.Rows[2];
                for (int index = 0; index < dt.Rows.Count; index++, ind++) {
                    r.Cells[1].Range.Text = ind.ToString();
                    r.Cells[2].Range.Text = dt.Rows[index].ItemArray[2].ToString();
                    r.Cells[3].Range.Text = "шт.";
                    r.Cells[4].Range.Text = dt.Rows[index].ItemArray[3].ToString();
                    count += int.Parse(dt.Rows[index].ItemArray[3].ToString());
                    r = tb.Rows.Add();
                }
                wordDocument.Range(r.Cells[1].Range.Start, r.Cells[2].Range.End).Cells.Merge();
                r.Cells[1].Range.Text = "ИТОГО";
                r.Cells[3].Range.Text = "{count}";
                ReplaceWordStub("{count}", count.ToString(), wordDocument);
                string fio = user.getSName() + " " + user.getName() + " " + user.getP();
                ReplaceWordStub("{empl}", fio, wordDocument);
                wordapp.Visible = true;
            } catch (Exception ex) {
                MessageBox.Show($"Произошла ошибка", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (wordDocument == null) { return; }
                wordDocument.Close();
            }
        }
        private void Form9_Shown(object sender, EventArgs e) {
            comboBox4.SelectedIndexChanged -= comboBox4_SelectedIndexChanged;
            sql = dataBase.getConnection();
            sql.Open();
            dateTimePicker1.MaxDate = DateTime.Today;
            dateTimePicker1.MinDate = DateTime.Today;
            dateTimePicker2.MinDate = DateTime.Today;
            dateTimePicker2.MaxDate = DateTime.Today.AddMonths(1);
            setMode();
            getAccess();
            sale();
            ComBx1();
            ComBx3();
            ComBx4();
            comboBox4.SelectedIndexChanged -= comboBox4_SelectedIndexChanged;
            DataTable dataTable = GetData($"select countAtStore from  products where id_product = {comboBox4.SelectedValue}");
            if (dataTable.Rows.Count == 0) {
                return;
            }
            numericUpDown1.Maximum = int.Parse(dataTable.Rows[0].ItemArray[0].ToString());
            numericUpDown1.Minimum = 1;
        }
        private void Form9_Load(object sender, EventArgs e) {
        }
        private void поСуммеToolStripMenuItem_Click(object sender, EventArgs e) {
            if (dataGridView1.Rows.Count == 0) {
                MessageBox.Show($"Сначала надо добавить данные о продаже", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            sql.Close();
            Form form = Form17.getInstance();
            form.ShowDialog();
        }
        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e) {
            DataTable dataTable = GetData($"select countAtStore from  products where id_product = {comboBox4.SelectedValue}");
            if (dataTable.Rows.Count == 0) {
                return;
            }
            numericUpDown1.Maximum = int.Parse(dataTable.Rows[0].ItemArray[0].ToString());
            numericUpDown1.Minimum = 1;
        }
        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e) {
        }
        private void label1_Click(object sender, EventArgs e) {
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e) {
        }
        private void label5_Click(object sender, EventArgs e) {
        }
        private void numericUpDown1_ValueChanged(object sender, EventArgs e) {
        }
        private void статистикаToolStripMenuItem_Click(object sender, EventArgs e) {
        }

        private void сформироватьТТНToolStripMenuItem_MouseHover(object sender, EventArgs e) {

        }

        private void cформироатьЗаявкуToolStripMenuItem_MouseHover(object sender, EventArgs e) {
        }

        private void статистикаToolStripMenuItem_MouseHover(object sender, EventArgs e) {
        }

        private void button7_Click(object sender, EventArgs e) {
            if (dataGridView2.Rows.Count == 0) {
                return;
            }
            if (dataGridView1.CurrentRow.Cells[7].Value.ToString() != "") {
                MessageBox.Show("Вы не можете удалить товары из уже доставленного заказа", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            try {
                string query = $"delete from saleTemp where id_pos = {dataGridView2.CurrentRow.Cells[0].Value}";
                SqlCommand sqlCommand = new SqlCommand(query, sql);
                sqlCommand.ExecuteNonQuery();
                MessageBox.Show("Данные успешно удалены", "SUCCESS", MessageBoxButtons.OK);
                saleTemp();
            } catch (Exception ex) {
                MessageBox.Show($"Не удалось удалить данные", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
