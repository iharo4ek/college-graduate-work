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
namespace Medplast {
    public partial class Form1 : Form {
        private static Form1 insatance;
        private Form1() {
            InitializeComponent();
        }
        User user = User.getInstance();
        DataBase dataBase = DataBase.getInstance();
        public static Form1 getInstance() {
            if (insatance == null) {
                insatance = new Form1();
            }
            return insatance;
        }
        private void Form1_Load(object sender, EventArgs e) {
            this.textBox2.AutoSize = false;
            this.textBox2.Size = new System.Drawing.Size(169, 25);
            textBox2.UseSystemPasswordChar = true;
            pictureBox2.Visible = false;
            textBox1.MaxLength = 30;
        }
        private void pictureBox3_Click(object sender, EventArgs e) {
            textBox2.UseSystemPasswordChar = false;
            pictureBox3.Visible = false;
            pictureBox2.Visible = true;
        }
        private void pictureBox2_Click(object sender, EventArgs e) {
            textBox2.UseSystemPasswordChar = true;
            pictureBox3.Visible = true;
            pictureBox2.Visible = false;
        }
        private void button1_Click(object sender, EventArgs e) {
            if (textBox1.Text == "" || textBox2.Text == "") {
                MessageBox.Show("Для входа необходимо ввести логин и пароль");
                return;
            }
            string loginUser = textBox1.Text;
            string passUser = textBox2.Text;
            SqlDataAdapter adapter = new SqlDataAdapter();
            DataTable table = new DataTable();
            string query = $"Select id_employee, employeeLogin, employeePassword, jobTitle, employeeSurname, employeeName, " +
                $"employeePatronymic from employees inner join jobTitles on employees.id_jobTitle = jobTitles.id_jobTitle " +
                $"where employeeLogin like N'{loginUser}' and employeePassword = N'{passUser}';";
            SqlCommand command = new SqlCommand(query, dataBase.getConnection());
            adapter.SelectCommand = command;
            adapter.Fill(table);
            if (table.Rows.Count == 1) {
                Form form = Form2.getInstance();
                user.setId(int.Parse(table.Rows[0].ItemArray[0].ToString()));
                user.setLogin(table.Rows[0].ItemArray[1].ToString());
                user.setPassword(table.Rows[0].ItemArray[2].ToString());
                user.setJobTitle(table.Rows[0].ItemArray[3].ToString());
                user.setSName(table.Rows[0].ItemArray[4].ToString());
                user.setName(table.Rows[0].ItemArray[5].ToString());
                user.setP(table.Rows[0].ItemArray[6].ToString());
                this.Hide();
                form.ShowDialog();
            } else {
                MessageBox.Show("не верный логин или пароль");
            }
        }
    }
}
