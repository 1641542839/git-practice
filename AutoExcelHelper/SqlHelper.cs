using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoExcelHelper
{
    class SqlHelper
    {
        public static DataTable ExecSQLSales(string sql_conn)
        {
            DataTable dt = new DataTable();
            //SqlConnection connection = new SqlConnection(Properties.Settings.Default.SalesHistoryConnectionString);
            //SqlCommand cmd = connection.CreateCommand();

            String sql = sql_conn.Split('|')[0];
            String connection = sql_conn.Split('|')[1];
            var conn = new SqlConnection(connection);
            // Console.WriteLine("connecting database...");

            try
            {

                SqlCommand cmd = new SqlCommand(sql, conn);

                // create data adapter
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                // store data to datatable
                da.Fill(dt);

                conn.Close();
                da.Dispose();

            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.WriteLine(e.Message);
            }

            return dt;
        }
    }
}
