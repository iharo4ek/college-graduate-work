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
    public partial class Form17 : Form {
        private Form17() {
            InitializeComponent();
        }
        DataBase dataBase = DataBase.getInstance();
        User user = User.getInstance();
        private static SqlConnection sql;
        private static Form instance;
        private static int q;
        public static void setQ(int i) {
            q = i;
        }
        public static Form getInstance() {
            if (instance == null) {
                instance = new Form17();
            }
            return instance;
        }
        public DataTable GetData(string query) {
            SqlDataAdapter adapter = new SqlDataAdapter(query, dataBase.getSrc());
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            return dt;
        }
        private void Form17_Load(object sender, EventArgs e) {
            sql = dataBase.getConnection();
            sql.Open();
            chart1.Series[0].Points.Clear();
            chart1.Series[0].IsVisibleInLegend = true;
            chart1.Series[0].Enabled = true;
            chart1.Series[0].Name = "Статистикса по сумме продаж";
            DataTable dt = GetData($"select departureDate,SUM(summ) from sale where departureDate <> '' group by departureDate");
            if (dt.Rows.Count == 0) { return; }
            for (int i = 0; i < dt.Rows.Count; i++) {
                DateTime date = DateTime.Parse(dt.Rows[i].ItemArray[0].ToString());
                chart1.Series[0].Points.AddXY(date.ToString("dd.MM.yyyy"), double.Parse(dt.Rows[i].ItemArray[1].ToString()));
            }
        }
        private void Form17_Shown(object sender, EventArgs e) {

            //if (q == 2) {
            //    chart1.Series[1].IsVisibleInLegend = true;
            //    chart1.Series[1].Enabled = true;
            //    chart1.Series[0].IsVisibleInLegend = false;
            //    chart1.Series[0].Enabled = false;
            //    chart1.Series[1].Name = "Статистикса по производству";
            //    DataTable dt = GetData($"select saleDate,SUM(summ) from sale group by saleDate");
            //    if (dt.Rows.Count == 0) { return; }
            //    double sum = 0;
            //    for (int i = 0; i < dt.Rows.Count; i++) {
            //        DateTime date = DateTime.Parse(dt.Rows[i].ItemArray[0].ToString());
            //        chart1.Series[0].Points.AddXY(date.ToString("yyyy-mm-dd"), double.Parse(dt.Rows[i].ItemArray[1].ToString()));
            //    }
            //}
        }
        private void Form17_FormClosing(object sender, FormClosingEventArgs e) {
            sql.Close();
        }
    }
}
