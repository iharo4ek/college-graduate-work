using System;
using System.Data.SqlClient;
namespace Medplast
{
    class DataBase
    {
        private string src = @"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=MedplastDataBase;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False";
        SqlConnection sqlconnection;
        private static DataBase instance;
        private DataBase() {
            sqlconnection = new SqlConnection(src);
        }
        public static DataBase getInstance() {
            if (instance == null) {
                instance = new DataBase();
            }
            return instance;
        }
        public SqlConnection getConnection()
        {
            return sqlconnection;
        }
        public String getSrc()
        {
            return src;
        }
    }
}
