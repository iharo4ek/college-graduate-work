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
    public partial class Form10 : Form {
        DataBase dataBase = DataBase.getInstance();
        User user = User.getInstance();
        private SqlConnection sql;
        private static Form instance;
        private static int q = -1;
        private Form10() {
            InitializeComponent();
        }
        public static Form getInstance() {
            if (instance == null) {
                instance = new Form10();
            }
            return instance;
        }
        private void Form10_Load(object sender, EventArgs e) {
            sql = dataBase.getConnection();
            sql.Open();
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
        private void Form10_FormClosing(object sender, FormClosingEventArgs e) {
            sql.Close();
        }
        public static void setQ(int i) {
            q = i;
        }
        private void button1_Click(object sender, EventArgs e) {
            if (dateTimePicker1.Value > dateTimePicker2.Value) {
                MessageBox.Show($"Дата начала  промежутка не может быть больше даты конца промежутка", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (q == 1) {
                var wordapp = new Word.Application();
                string path = Environment.CurrentDirectory + @"\Otchet.docx";
                var wordDocument = wordapp.Documents.Open(path);
                try {
                    wordapp.Visible = false;
                    DataTable dt = GetData($"select id_sale, nameOrganisation, summ from sale inner join  clients on sale.id_client = clients.id_client where saleDate >= '{dateTimePicker1.Value.ToString("yyyy-MM-dd")}' and saleDate <= '{dateTimePicker2.Value.ToString("yyyy-MM-dd")}';");
                    ReplaceWordStub("{date1}", dateTimePicker1.Value.ToString("yyyy-MM-dd"), wordDocument);
                    ReplaceWordStub("{date2}", dateTimePicker2.Value.ToString("yyyy-MM-dd"), wordDocument);
                    Word.Table tb = wordDocument.Tables[1];
                    Word.Row r = tb.Rows[2];
                    double sum = 0;
                    for (int index = 0; index < dt.Rows.Count; index++) {
                        r.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray625;
                        r.Cells[1].Range.Text = dt.Rows[index].ItemArray[1].ToString();
                        r.Cells[4].Range.Text = dt.Rows[index].ItemArray[2].ToString();
                        DataTable dt2 = GetData($"select (productName + N' шт.') as [товар], countProducts, cost from sale inner join saleTemp on sale.id_sale = saleTemp.id_sale inner join products on saleTemp.id_product = products.id_product where sale.id_sale = {int.Parse(dt.Rows[index].ItemArray[0].ToString())} and  saleDate >= '{dateTimePicker1.Value.ToString("yyyy-MM-dd")}' and saleDate <= '{dateTimePicker2.Value.ToString("yyyy-MM-dd")}';");
                        sum += double.Parse(dt.Rows[index].ItemArray[2].ToString());
                        r = tb.Rows.Add();
                        r.Shading.BackgroundPatternColor = Word.WdColor.wdColorWhite;
                        for (int index2 = 0; index2 < dt2.Rows.Count; index2++) {
                            r.Cells[1].Range.Text = dt2.Rows[index2].ItemArray[0].ToString();
                            r.Cells[2].Range.Text = dt2.Rows[index2].ItemArray[1].ToString();
                            r.Cells[3].Range.Text = dt2.Rows[index2].ItemArray[2].ToString();
                            double s = double.Parse(dt2.Rows[index2].ItemArray[1].ToString()) * double.Parse(dt2.Rows[index2].ItemArray[2].ToString());
                            r.Cells[4].Range.Text = s.ToString();
                            r = tb.Rows.Add();
                        }
                    }
                    r.Cells[1].Range.Text = "ИТОГО";
                    r.Cells[4].Range.Text = "{sum}";
                    ReplaceWordStub("{sum}", sum.ToString(), wordDocument);
                    string fio = user.getSName() + " " + user.getName() + " " + user.getP();
                    ReplaceWordStub("{empl}", fio, wordDocument);
                    wordapp.Visible = true;
                } catch (Exception ex) {
                    MessageBox.Show($"Произошла ошибка", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    if (wordDocument == null) { return; }
                    wordDocument.Close();
                }
            }
            if (q == 2) {
                var wordapp = new Word.Application();
                string path = Environment.CurrentDirectory + @"\O1.docx";
                var wordDocument = wordapp.Documents.Open(path);
                try {
                    wordapp.Visible = false;
                    DataTable dt = GetData($"select (select (employeeSurname + ' ' + employeeName + ' ' + employeePatronymic) from employees where id_employee = id_manager) ,COUNT(id_manager), round(sum(summ),2), COUNT(id_pos) from sale inner join saleTemp on sale.id_sale = saleTemp.id_sale where saleDate >= '{dateTimePicker1.Value.ToString("yyyy-MM-dd")}' and saleDate <= '{dateTimePicker2.Value.ToString("yyyy-MM-dd")}'  group by id_manager;");
                    ReplaceWordStub("{date1}", dateTimePicker1.Value.ToString("yyyy-MM-dd"), wordDocument);
                    ReplaceWordStub("{date2}", dateTimePicker2.Value.ToString("yyyy-MM-dd"), wordDocument);
                    Word.Table tb = wordDocument.Tables[1];
                    Word.Row r = tb.Rows[2];
                    int count1 = 0, count2 = 0;
                    double sum = 0;
                    for (int index = 0; index < dt.Rows.Count; index++) {
                        r.Cells[1].Range.Text = dt.Rows[index].ItemArray[0].ToString();
                        r.Cells[2].Range.Text = dt.Rows[index].ItemArray[1].ToString();
                        r.Cells[3].Range.Text = dt.Rows[index].ItemArray[2].ToString();
                        r.Cells[4].Range.Text = dt.Rows[index].ItemArray[3].ToString();
                        count1 += int.Parse(dt.Rows[index].ItemArray[1].ToString());
                        count2 += int.Parse(dt.Rows[index].ItemArray[3].ToString());
                        sum += double.Parse(dt.Rows[index].ItemArray[2].ToString());
                        r = tb.Rows.Add();
                    }
                    r.Cells[1].Range.Text = "ИТОГО";
                    r.Cells[2].Range.Text = count1.ToString();
                    r.Cells[3].Range.Text = sum.ToString();
                    r.Cells[4].Range.Text = count2.ToString();
                    wordapp.Visible = true;
                } catch (Exception ex) {
                    MessageBox.Show($"Произошла ошибка", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    if (wordDocument == null) { return; }
                    wordDocument.Close();
                }
            }
            if (q == 3) {
                var wordapp = new Word.Application();
                string path = Environment.CurrentDirectory + @"\O2.docx";
                var wordDocument = wordapp.Documents.Open(path);
                try {

                    wordapp.Visible = false;
                    DataTable dt = GetData($"select nameOrganisation, COUNT(nameOrganisation), round(SUM(summ),2), COUNT(id_pos) from sale inner join saleTemp on sale.id_sale = saleTemp.id_sale inner join clients on sale.id_client = clients.id_client where saleDate >= '{dateTimePicker1.Value.ToString("yyyy-MM-dd")}' and saleDate <= '{dateTimePicker2.Value.ToString("yyyy-MM-dd")}' group by nameOrganisation;");
                    ReplaceWordStub("{date1}", dateTimePicker1.Value.ToString("yyyy-MM-dd"), wordDocument);
                    ReplaceWordStub("{date2}", dateTimePicker2.Value.ToString("yyyy-MM-dd"), wordDocument);
                    Word.Table tb = wordDocument.Tables[1];
                    Word.Row r = tb.Rows[2];
                    int count1 = 0, count2 = 0;
                    double sum = 0;
                    for (int index = 0; index < dt.Rows.Count; index++) {
                        r.Cells[1].Range.Text = dt.Rows[index].ItemArray[0].ToString();
                        r.Cells[2].Range.Text = dt.Rows[index].ItemArray[1].ToString();
                        r.Cells[3].Range.Text = dt.Rows[index].ItemArray[2].ToString();
                        r.Cells[4].Range.Text = dt.Rows[index].ItemArray[3].ToString();
                        count1 += int.Parse(dt.Rows[index].ItemArray[1].ToString());
                        count2 += int.Parse(dt.Rows[index].ItemArray[3].ToString());
                        sum += double.Parse(dt.Rows[index].ItemArray[2].ToString());
                        r = tb.Rows.Add();
                    }
                    r.Cells[1].Range.Text = "ИТОГО";
                    r.Cells[2].Range.Text = count1.ToString();
                    r.Cells[3].Range.Text = sum.ToString();
                    r.Cells[4].Range.Text = count2.ToString();
                    wordapp.Visible = true;
                } catch (Exception ex) {
                    MessageBox.Show($"Произошла ошибка", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    if (wordDocument == null) { return; }
                    wordDocument.Close();
                }
            }
            if (q == 4) {
                var wordapp = new Word.Application();
                string path = Environment.CurrentDirectory + @"\O3.docx";
                var wordDocument = wordapp.Documents.Open(path);
                try {

                    wordapp.Visible = false;
                    DataTable dt = GetData($"select productName, SUM(countProducts), round(sum(countProducts*cost),2) from sale inner join saleTemp on sale.id_sale = saleTemp.id_sale inner join products on saleTemp.id_product = products.id_product where saleDate >= '{dateTimePicker1.Value.ToString("yyyy-MM-dd")}' and saleDate <= '{dateTimePicker2.Value.ToString("yyyy-MM-dd")}' group by productName;");
                    ReplaceWordStub("{date1}", dateTimePicker1.Value.ToString("yyyy-MM-dd"), wordDocument);
                    ReplaceWordStub("{date2}", dateTimePicker2.Value.ToString("yyyy-MM-dd"), wordDocument);
                    Word.Table tb = wordDocument.Tables[1];
                    Word.Row r = tb.Rows[2];
                    double sum = 0; 
                    int count = 0;
                    for (int index = 0; index < dt.Rows.Count; index++) {
                        r.Cells[1].Range.Text = dt.Rows[index].ItemArray[0].ToString();
                        r.Cells[2].Range.Text = dt.Rows[index].ItemArray[1].ToString();
                        r.Cells[3].Range.Text = dt.Rows[index].ItemArray[2].ToString();
                        count += int.Parse(dt.Rows[index].ItemArray[1].ToString());
                        sum += double.Parse(dt.Rows[index].ItemArray[2].ToString());
                        r = tb.Rows.Add();
                    }
                    r.Cells[1].Range.Text = "ИТОГО";
                    r.Cells[3].Range.Text = sum.ToString();
                    r.Cells[2].Range.Text = count.ToString();
                    wordapp.Visible = true;
                } catch {
                    MessageBox.Show($"Произошла ошибка", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    if (wordDocument == null) { return; }
                    wordDocument.Close();
                }
            }
            this.Close();
        }
    }
}
