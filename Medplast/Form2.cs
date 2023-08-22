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
    public partial class Form2 : Form {
        private static Form instance;
        Form form;
        private SqlConnection sql;
        private DataBase dataBase = DataBase.getInstance();
        private Form2() {
            InitializeComponent();
        }
        User user = User.getInstance();
        public static Form getInstance() {
            if (instance == null)
                instance = new Form2();
            return instance;
        }
        public void getAccess() {
            switch (user.getJobTitle()) {
                case "бухгалтер": {
                        button1.Enabled = true;
                        button2.Enabled = true;
                        button3.Enabled = true;
                        button4.Enabled = true;
                        button5.Enabled = true;
                        button7.Enabled = true;
                        button8.Enabled = true;
                        button9.Enabled = true;
                        button10.Enabled = true;
                        button11.Enabled = true;
                        button12.Enabled = true;
                        button13.Enabled = true;
                        break;
                    }
                case "главный бухгалтер": {
                        button1.Enabled = true;
                        button2.Enabled = true;
                        button3.Enabled = true;
                        button4.Enabled = true;
                        button5.Enabled = true;
                        button7.Enabled = true;
                        button8.Enabled = true;
                        button9.Enabled = true;
                        button10.Enabled = true;
                        button11.Enabled = true;
                        button12.Enabled = true;
                        button13.Enabled = true;
                        break;
                    }
                case "директор": {
                        button1.Enabled = true;
                        button2.Enabled = true;
                        button3.Enabled = true;
                        button4.Enabled = true;
                        button5.Enabled = true;
                        button7.Enabled = true;
                        button8.Enabled = true;
                        button9.Enabled = true;
                        button10.Enabled = true;
                        button11.Enabled = true;
                        button12.Enabled = true;
                        button13.Enabled = true;
                        break;
                    }
                case "зам директора": {
                        button1.Enabled = true;
                        button2.Enabled = true;
                        button3.Enabled = true;
                        button4.Enabled = true;
                        button5.Enabled = true;
                        button7.Enabled = true;
                        button8.Enabled = true;
                        button9.Enabled = true;
                        button10.Enabled = true;
                        button11.Enabled = true;
                        button12.Enabled = true;
                        button13.Enabled = true;
                        break;
                    }
                case "инженер": {
                        button8.Enabled = true;
                        button9.Enabled = true;
                        button10.Enabled = true;
                        break;
                    }
                case "мастер цеха": {
                        button11.Enabled = true;
                        button12.Enabled = true;
                        button13.Enabled = true;
                        break;
                    }
                case "менеджер": {
                        button7.Enabled = true;
                        button5.Enabled = true;
                        button11.Enabled = true;
                        break;
                    }
                case "администратор": {
                        button1.Enabled = true;
                        button2.Enabled = true;
                        button3.Enabled = true;
                        button4.Enabled = true;
                        button5.Enabled = true;
                        button7.Enabled = true;
                        button8.Enabled = true;
                        button9.Enabled = true;
                        button10.Enabled = true;
                        button11.Enabled = true;
                        button12.Enabled = true;
                        button13.Enabled = true;
                        break;
                    }
                default: {
                        button1.Enabled = true;
                        button2.Enabled = true;
                        button3.Enabled = true;
                        button4.Enabled = true;
                        button5.Enabled = true;
                        button7.Enabled = true;
                        button8.Enabled = true;
                        button9.Enabled = true;
                        button10.Enabled = true;
                        button11.Enabled = true;
                        button12.Enabled = true;
                        button13.Enabled = true;
                        break;
                    }
            }
        }
        private void Form2_Load(object sender, EventArgs e) {
            getAccess();
        }
        private void button1_Click(object sender, EventArgs e) {
            form = Form3.getInstance();
            this.Hide();
            form.ShowDialog();
        }
        private void Form2_FormClosing(object sender, FormClosingEventArgs e) {
            Form form = Form1.getInstance();
            form.Close();
        }
        private void button2_Click(object sender, EventArgs e) {
            form = Form4.getInstance();
            this.Hide();
            form.ShowDialog();
        }
        private void button4_Click(object sender, EventArgs e) {
            form = Form6.getInstance();
            this.Hide();
            form.ShowDialog();
        }
        private void button3_Click(object sender, EventArgs e) {
            form = Form5.getInstance();
            this.Hide();
            form.ShowDialog();
        }
        private void button7_Click(object sender, EventArgs e) {
            form = Form7.getInstance();
            this.Hide();
            form.ShowDialog();
        }
        private void button8_Click(object sender, EventArgs e) {
            form = Form8.getInstance();
            this.Hide();
            form.ShowDialog();
        }
        private void button5_Click(object sender, EventArgs e) {
            form = Form9.getInstance();
            this.Hide();
            form.ShowDialog();
        }
        private void button10_Click(object sender, EventArgs e) {
            form = Form11.getInstance();
            this.Hide();
            form.ShowDialog();
        }
        private void button9_Click(object sender, EventArgs e) {
            form = Form12.getInstance();
            this.Hide();
            form.ShowDialog();
        }
        private void button12_Click(object sender, EventArgs e) {
            form = Form13.getInstance();
            this.Hide();
            form.ShowDialog();
        }
        private void button11_Click(object sender, EventArgs e) {
            form = Form14.getInstance();
            this.Hide();
            form.ShowDialog();
        }
        private void button13_Click(object sender, EventArgs e) {
            form = Form15.getInstance();
            this.Hide();
            form.ShowDialog();
        }
        private void ReplaceWordStub(string stubToReplace, string text, Word.Document wordDocumet) {
            var range = wordDocumet.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: stubToReplace, ReplaceWith: text);
        }
        private void button6_Click(object sender, EventArgs e) {
            Form form = Form10.getInstance();
            Form10.setQ(1);
            form.ShowDialog();
        }
        public DataTable GetData(string query) {
            SqlDataAdapter adapter = new SqlDataAdapter(query, dataBase.getSrc());
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            return dt;
        }
        private void button14_Click(object sender, EventArgs e) {
        }
        private void Form2_Shown(object sender, EventArgs e) {
        }
        private void button14_Click_1(object sender, EventArgs e) {
            Form form = Form10.getInstance();
            Form10.setQ(2);
            form.ShowDialog();
        }
        private void button15_Click(object sender, EventArgs e) {
            Form form = Form10.getInstance();
            Form10.setQ(3);
            form.ShowDialog();
        }
        private void button16_Click(object sender, EventArgs e) {
            Form form = Form10.getInstance();
            Form10.setQ(4);
            form.ShowDialog();
        }
    }
}
