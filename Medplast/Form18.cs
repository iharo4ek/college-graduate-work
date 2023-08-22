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
    public partial class Form18 : Form {
        private DataBase dataBase = DataBase.getInstance();
        private User user;
        private SqlConnection sql;
        private Form18() {
            InitializeComponent();
        }
        private static Form instance;
        public static Form getInstance() {
            if (instance == null) {
                instance = new Form18();
            }
            return instance;
        }
        private DataTable GetData(string query) {
            SqlDataAdapter adapter = new SqlDataAdapter(query, dataBase.getSrc());
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            return dt;
        }
        private void Form18_Shown(object sender, EventArgs e) {
            sql = dataBase.getConnection();
            sql.Open();
            this.Text = "список недобросовестных поставщиков";
            string q = $"select nameOrganisation as [поставщик],purchaseDate as [дата заказа], receivingDate as [дата получения], DATEDIFF(DAY,purchaseDate,receivingDate) as [разница]  from materialsPurchase inner join providers on materialsPurchase.id_provider = providers.id_provider" +
                $" where (DATEDIFF(DAY,purchaseDate,receivingDate) > 30 or ('{DateTime.Now.ToString("yyyy-MM-dd")}'>DateAdd(day,30,purchaseDate) and receivingDate is null)) and DATEDIFF(MONTH, purchaseDate, GETDATE()) <= 6";
            DataTable dt = GetData(q);
            dataGridView1.DataSource = dt;
            dataGridView1.AutoResizeColumns();
        }
        private void Form18_FormClosing(object sender, FormClosingEventArgs e) {
            sql.Close();
        }

        private void Form18_Load(object sender, EventArgs e) {

        }
    }
}
