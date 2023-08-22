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
using Word = Microsoft.Office.Interop.Word;
namespace Medplast {
    public partial class Form12 : Form {
        private Modes mode;
        private const string STRING_PATTERN = @"[^а-яА-яa-zA-Z]";
        private const string INT_PATTERN = @"[^0-9.]";
        DataBase dataBase = DataBase.getInstance();
        SqlConnection sql;
        User user = User.getInstance();
        private Form12() {
            InitializeComponent();
        }
        private static Form instance;
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
                button1.Visible = false;
                button2.Visible = false;
                button3.Visible = false;
                button6.Visible = false;
                comboBox1.Visible = false;
                comboBox2.Visible = false;
                comboBox3.Visible = false;
                dateTimePicker2.Visible = false;
                numericUpDown1.Visible = false;
                numericUpDown2.Visible = false;
                label1.Visible = false;
                label2.Visible = false;
                label3.Visible = false;
                label4.Visible = false;
                label5.Visible = false;
                label6.Visible = false;
            }
        }
        public static Form getInstance() {
            if (instance == null) {
                instance = new Form12();
            }
            return instance;
        }
        private void purchase() {
            string query = "select id_purchase, purchaseDate as [дата покупки], round(summ,2) as [сумма заказа, бел. руб], numberOfTheCar as  [гос. номер], nameOrganisation as [поставщик], receivingDate as [дата получения] from materialsPurchase left join cars on materialsPurchase.id_car = cars.id_car inner join providers on materialsPurchase.id_provider = providers.id_provider;";
            SqlDataAdapter adapter = new SqlDataAdapter(query, dataBase.getSrc());
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            dataGridView1.DataSource = dt;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.AutoResizeColumns();
        }
        private void purchaseTemp() {
            try {
                if (dataGridView1.CurrentRow == null) { return; }
                string query = $"select id_pos, id_purchase, materialName as [материал], countMaterials as [количество], cost as [цена за ед, бел. руб.], unit as [ед.изм.]  from purchaseTemp inner join materials on purchaseTemp.id_material = materials.id_material where id_purchase = {dataGridView1.CurrentRow.Cells[0].Value}";
                SqlDataAdapter adapter = new SqlDataAdapter(query, dataBase.getSrc());
                DataTable dt = new DataTable();
                adapter.Fill(dt);
                dataGridView2.DataSource = dt;
                dataGridView2.Columns[0].Visible = false;
                dataGridView2.Columns[1].Visible = false;
                dataGridView2.AutoResizeColumns();
            } catch {
            }
        }
        void ComBx1() {
            string query = "select id_car, numberOfTheCar from cars;";
            SqlDataAdapter adapter = new SqlDataAdapter(query, dataBase.getSrc());
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            comboBox1.DataSource = dt;
            comboBox1.ValueMember = "id_car";
            comboBox1.DisplayMember = "numberOfTheCar";
        }
        void ComBx2() {
            string query = "select id_provider, nameOrganisation from providers;";
            SqlDataAdapter adapter = new SqlDataAdapter(query, dataBase.getSrc());
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            comboBox2.DataSource = dt;
            comboBox2.ValueMember = "id_provider";
            comboBox2.DisplayMember = "nameOrganisation";
        }
        void ComBx3() {
            string query = "select id_material, (materialName + ' ' + STR(countAtStore) + unit) as  [mat] from materials;";
            SqlDataAdapter adapter = new SqlDataAdapter(query, dataBase.getSrc());
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            comboBox3.DataSource = dt;
            comboBox3.ValueMember = "id_material";
            comboBox3.DisplayMember = "mat";
        }
        private void Form12_Load(object sender, EventArgs e) { }
        private void Form12_FormClosing(object sender, FormClosingEventArgs e) {
            sql.Close();
            Form form = Form2.getInstance();
            form.Show();
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
        private void button2_Click(object sender, EventArgs e) {
            if (dataGridView1.Rows.Count == 0) { return; }
            try {
                string query = $"Delete From materialsPurchase Where id_purchase = {dataGridView1.CurrentRow.Cells[0].Value}";
                SqlCommand sqlCommand = new SqlCommand(query, sql);
                sqlCommand.ExecuteNonQuery();
                MessageBox.Show("Данные успешно удалены", "SUCCESS", MessageBoxButtons.OK);
                purchase();
                purchaseTemp();
                if (dataGridView1.Rows.Count == 0) {
                    dataGridView2.DataSource = null;
                }
            } catch (Exception ex) {
                MessageBox.Show($"Не удалось удалить данные Произошла ошибка", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void button1_Click(object sender, EventArgs e) {
            try {
                string query = $"Insert into materialsPurchase values ('{dateTimePicker1.Value.ToString("yyyy-MM-dd")}',0, null, {comboBox2.SelectedValue}, null);";
                SqlCommand sqlCommand = new SqlCommand(query, sql);
                sqlCommand.ExecuteNonQuery();
                MessageBox.Show("Данные успешно добавлены", "SUCCESS", MessageBoxButtons.OK);
                purchase();
            } catch (Exception ex) {
                MessageBox.Show($"{ex.Message}", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void button6_Click(object sender, EventArgs e) {
            try {
                if (dataGridView1.CurrentRow.Cells[5].Value.ToString() != "") {
                    MessageBox.Show("Данный заказ уже доставлен", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (numericUpDown2.Value <= 0) {
                    MessageBox.Show("Количество материалов должно быть больше 0", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (numericUpDown1.Value <= 0) {
                    MessageBox.Show("Цена материалов должна быть больше 0", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                DataTable dt = GetData($"select * from purchaseTemp where id_material = {comboBox3.SelectedValue} and id_purchase = {dataGridView1.CurrentRow.Cells[0].Value}");
                if (dt.Rows.Count != 0) {
                    MessageBox.Show("Данный материал уже имеется в заказе", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                string s = numericUpDown1.Value.ToString().Replace(',', '.'), s2 = numericUpDown2.Value.ToString().Replace(',', '.');
                string query = $"Insert into purchaseTemp values ({dataGridView1.CurrentRow.Cells[0].Value}, {comboBox3.SelectedValue}, " + s2 + ", " + s + ");";
                SqlCommand sqlCommand = new SqlCommand(query, sql);
                sqlCommand.ExecuteNonQuery();
                MessageBox.Show("Данные успешно добавлены", "SUCCESS", MessageBoxButtons.OK);
                purchase();
                purchaseTemp();
            } catch (Exception ex) {
                MessageBox.Show($"{ex.Message}", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e) {
            comboBox1.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            comboBox2.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
            button6.Enabled = true;
            comboBox3.Enabled = true;
            numericUpDown1.Enabled = true;
            numericUpDown2.Enabled = true;
            dateTimePicker2.MaxDate = DateTime.Parse(dataGridView1.CurrentRow.Cells[1].Value.ToString()).AddYears(1);
            purchaseTemp();
        }
        private void execHp() {
            SqlConnection con = new SqlConnection(dataBase.getSrc());
            SqlCommand cmd = new SqlCommand("AddMaterials", con);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@id_Purchase", dataGridView1.CurrentRow.Cells[0].Value);
            con.Open();
            int rowAffected = cmd.ExecuteNonQuery();
            con.Close();
        }
        private void button3_Click(object sender, EventArgs e) {
            try {
                if (dataGridView1.Rows.Count == 0) { return; }
                DataTable dataTable = GetData($"select * from materialsPurchase where id_car = {comboBox1.SelectedValue} and receivingDate = '{dateTimePicker2.Value.ToString("yyyy-MM-dd")}'");
                if (dataTable.Rows.Count != 0) {
                    MessageBox.Show($"Данная машиназанята в этот день", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                dataTable = GetData($"select * from sale where id_car = {comboBox1.SelectedValue} and departureDate = '{dateTimePicker2.Value.ToString("yyyy-MM-dd")}'");
                if (dataTable.Rows.Count != 0) {
                    MessageBox.Show($"Данная машиназанята в этот день", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (dataGridView2.Rows.Count == 0) {
                    MessageBox.Show($"Для начала нужно добавить товары в заказ", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                execHp();
                string query = $"Update materialsPurchase Set id_car = {comboBox1.SelectedValue}, " +
                    $"id_provider =  {comboBox2.SelectedValue}, receivingDate = " +
                    $"'{dateTimePicker2.Value.ToString("yyyy-MM-dd")}'  where id_purchase = {dataGridView1.CurrentRow.Cells[0].Value}";
                SqlCommand com = new SqlCommand(query, sql);
                com.ExecuteNonQuery();
                ComBx3();
                purchase();
                MessageBox.Show("Данные успешно изменены", "SUCCESS", MessageBoxButtons.OK);
            } catch (Exception ex) {
                MessageBox.Show($"Не удалось изменить данные Произошла ошибка", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
        private void сформироватьАктПриемаToolStripMenuItem_Click(object sender, EventArgs e) {
            if (dataGridView1.CurrentRow.Cells[5].Value.ToString() == "") {
                MessageBox.Show("Нельзя сформировать акт приема для неполученных материалов", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            var wordapp = new Word.Application();
            string path = Environment.CurrentDirectory + @"\Priem.docx";
            var wordDocument = wordapp.Documents.Open(path);
            try {
                wordapp.Visible = false;
                DataTable dt = GetData($"select (nameOrganisation + ' ' + adress) as " +
                    $"[provider],adress,materialName, countMaterials, cost, summ from " +
                    $"materialsPurchase inner join purchaseTemp on materialsPurchase.id_purchase = " +
                    $"purchaseTemp.id_purchase inner join materials on purchaseTemp.id_material = " +
                    $"materials.id_material inner join providers on materialsPurchase.id_provider = " +
                    $"providers.id_provider where materialsPurchase.id_purchase = {dataGridView1.CurrentRow.Cells[0].Value};");
                ReplaceWordStub("{provider}", dt.Rows[0].ItemArray[0].ToString(), wordDocument);
                ReplaceWordStub("{adress}", dt.Rows[0].ItemArray[1].ToString(), wordDocument);
                Word.Table tb = wordDocument.Tables[1];
                double stoim = 0; double sumNDS = 0; double stoimNDS = 0; double count = 0;
                Word.Row r = tb.Rows[3];
                for (int index = 0; index < dt.Rows.Count; index++) {
                    double currSumNds = 0, currStoimNds = 0;
                    currSumNds = double.Parse(dt.Rows[index].ItemArray[5].ToString()) * 0.2;
                    currStoimNds = currSumNds + double.Parse(dt.Rows[index].ItemArray[5].ToString());
                    r.Cells[1].Range.Text = dt.Rows[index].ItemArray[2].ToString();
                    r.Cells[2].Range.Text = "кг";
                    r.Cells[3].Range.Text = dt.Rows[index].ItemArray[3].ToString();
                    r.Cells[4].Range.Text = dt.Rows[index].ItemArray[4].ToString();
                    r.Cells[5].Range.Text = dt.Rows[index].ItemArray[5].ToString();
                    r.Cells[6].Range.Text = "20%";
                    r.Cells[7].Range.Text = currSumNds.ToString();
                    r.Cells[8].Range.Text = currStoimNds.ToString();
                    count += double.Parse(dt.Rows[index].ItemArray[3].ToString());
                    stoim += double.Parse(dt.Rows[index].ItemArray[5].ToString());
                    sumNDS += currSumNds; stoimNDS += currStoimNds;
                    r = tb.Rows.Add();
                }
                r.Cells[1].Range.Text = "ИТОГО"; r.Cells[2].Range.Text = "x"; r.Cells[3].Range.Text = "{c}";
                r.Cells[4].Range.Text = "x"; r.Cells[5].Range.Text = "{st}"; r.Cells[6].Range.Text = "x";
                r.Cells[7].Range.Text = "{sum}"; r.Cells[8].Range.Text = "{st}";
                ReplaceWordStub("{c}", count.ToString(), wordDocument); ReplaceWordStub("{st}", stoim.ToString(), wordDocument);
                ReplaceWordStub("{sum}", sumNDS.ToString(), wordDocument); ReplaceWordStub("{sum2}", sumNDS.ToString(), wordDocument);
                ReplaceWordStub("{st}", stoimNDS.ToString(), wordDocument); wordapp.Visible = true;
                ReplaceWordStub("{st2}", stoimNDS.ToString(), wordDocument); wordapp.Visible = true;
            } catch (Exception ex) {
                MessageBox.Show($"Произошла ошибка", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (wordDocument == null) { return; }
                wordDocument.Close();
            }
        }
        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e) {
            comboBox3.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            numericUpDown2.Text = dataGridView2.CurrentRow.Cells[3].Value.ToString();
            numericUpDown1.Text = dataGridView2.CurrentRow.Cells[4].Value.ToString();
        }
        private void сформироватьЗаказToolStripMenuItem_Click(object sender, EventArgs e) {
            if (dataGridView1.Rows.Count == 0) {
                MessageBox.Show($"Сначала нужно добавить данные о покупке материалов", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (dataGridView2.Rows.Count == 0) {
                MessageBox.Show($"Сначала нужно добавить данные о составе покупки материалов", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            var wordapp = new Word.Application();
            string path = Environment.CurrentDirectory + @"\Zakaz.docx";
            var wordDocument = wordapp.Documents.Open(path);
            try {
                int ind = 1;
                wordapp.Visible = false;
                DataTable dt = GetData($"select purchaseDate, nameOrganisation, materialName, countMaterials, cost, unit from materialsPurchase inner join purchaseTemp on materialsPurchase.id_purchase = purchaseTemp.id_purchase inner join providers on materialsPurchase.id_provider = providers.id_provider inner join materials on purchaseTemp.id_material = materials.id_material where materialsPurchase.id_purchase = {dataGridView1.CurrentRow.Cells[0].Value}");
                DateTime dateTime = DateTime.Parse(dt.Rows[0].ItemArray[0].ToString());
                ReplaceWordStub("{date}", dateTime.ToString("dd.MM.yyyy"), wordDocument);
                ReplaceWordStub("{provider}", dt.Rows[0].ItemArray[1].ToString(), wordDocument);
                Word.Table tb = wordDocument.Tables[1];
                Word.Row r = tb.Rows[2];
                double count = 0, sum = 0;
                for (int index = 0; index < dt.Rows.Count; index++, ind++) {
                    r.Cells[1].Range.Text = ind.ToString();
                    r.Cells[2].Range.Text = dt.Rows[index].ItemArray[2].ToString();
                    r.Cells[3].Range.Text = dt.Rows[index].ItemArray[5].ToString();
                    r.Cells[4].Range.Text = dt.Rows[index].ItemArray[4].ToString();
                    r.Cells[5].Range.Text = dt.Rows[index].ItemArray[3].ToString();
                    sum += double.Parse(dt.Rows[index].ItemArray[4].ToString());
                    count += double.Parse(dt.Rows[index].ItemArray[3].ToString());
                    r = tb.Rows.Add();
                }
                wordDocument.Range(r.Cells[1].Range.Start, r.Cells[2].Range.End).Cells.Merge();
                r.Cells[1].Range.Text = "ИТОГО";
                r.Cells[3].Range.Text = "{sum}";
                r.Cells[4].Range.Text = "{count}";
                ReplaceWordStub("{sum}", sum.ToString(), wordDocument);
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
        private void Form12_Shown(object sender, EventArgs e) {
            numericUpDown1.DecimalPlaces = 2;
            numericUpDown1.Increment = 0.1M;
            numericUpDown2.DecimalPlaces = 2;
            numericUpDown2.Increment = 0.1M;
            numericUpDown2.Maximum = 20000;
            numericUpDown2.Minimum = 50;
            numericUpDown1.Maximum = 100;
            numericUpDown2.Minimum = 0.1M;
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("en-US");
            sql = dataBase.getConnection();
            sql.Open();
            dateTimePicker1.MaxDate = DateTime.Today;
            dateTimePicker1.MinDate = DateTime.Today;
            dateTimePicker2.MinDate = DateTime.Today;
            setMode();
            getAccess();
            purchase();
            ComBx1();
            ComBx2();
            ComBx3();
        }
        private void списокНеДобросовестныхПоставщиковToolStripMenuItem_Click(object sender, EventArgs e) {
            if (dataGridView1.RowCount == 0) {
                return;
            }
            sql.Close();
            Form form = Form18.getInstance();
            form.ShowDialog();
            sql.Open();
        }

        private void сформироватьАктПриемаToolStripMenuItem_MouseHover(object sender, EventArgs e) {
            this.BackColor = Color.Teal;
        }

        private void сформироватьЗаказToolStripMenuItem_MouseHover(object sender, EventArgs e) {
            this.BackColor = Color.Teal;
        }

        private void списокНеДобросовестныхПоставщиковToolStripMenuItem_MouseHover(object sender, EventArgs e) {
            this.BackColor = Color.Teal;
        }

        private void button7_Click(object sender, EventArgs e) {
            if (dataGridView2.Rows.Count == 0) { return; }
            if (dataGridView1.CurrentRow.Cells[5].Value.ToString() != "") {
                MessageBox.Show("Данный заказ уже доставлен", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            try {
                string query = $"delete from purchaseTemp where id_pos = {dataGridView2.CurrentRow.Cells[0].Value}";
                SqlCommand sqlCommand = new SqlCommand(query, sql);
                sqlCommand.ExecuteNonQuery();
                MessageBox.Show("Данные успешно удалены", "SUCCESS", MessageBoxButtons.OK);
                purchaseTemp();
            } catch (Exception ex) {
                MessageBox.Show($"Не удалось удалить данные Произошла ошибка", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e) {
            
        }

        private void comboBox3_TextChanged(object sender, EventArgs e) {
            
        }
    }
}
