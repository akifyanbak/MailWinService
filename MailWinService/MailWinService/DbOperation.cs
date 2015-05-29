using System.Configuration;
using System.Data;
using System.Data.SqlClient;

namespace MailWinService
{
    public class DbOperation
    {
        static readonly string Constr = ConfigurationManager.ConnectionStrings["MyDBConnectionString"].ConnectionString;

        public static DataTable Data(string[] parameters, string[] values, string command)
        {
            var connection = new SqlConnection(Constr);
            var sqlCommand = new SqlCommand(command, connection);
            for (int i = 0; i < parameters.Length; i++)
            {
                sqlCommand.Parameters.AddWithValue(parameters[i], values[i]);
            }
            var dataAdapter = new SqlDataAdapter(sqlCommand);
            var dataTable = new DataTable();
            dataAdapter.Fill(dataTable);

            if (connection.State == ConnectionState.Open) { connection.Close(); }

            return dataTable;
        }
        public static DataTable Data(string command)
        {
            var connection = new SqlConnection(Constr);
            connection.Open();
            var sqlCommand = new SqlCommand(command, connection);
            var dataAdapter = new SqlDataAdapter(sqlCommand);
            var dataTable = new DataTable();
            dataAdapter.Fill(dataTable);

            if (connection.State == ConnectionState.Open) { connection.Close(); }

            return dataTable;
        }
        public static int Execute(string[] parameters, string[] values, string command)
        {
            var connection = new SqlConnection(Constr);
            var sqlCommand = new SqlCommand(command, connection);
            for (int i = 0; i < parameters.Length; i++)
            {
                sqlCommand.Parameters.AddWithValue(parameters[i], values[i]);
            }

            if (connection.State != ConnectionState.Open) { connection.Open(); }

            var result = sqlCommand.ExecuteNonQuery();

            if (connection.State == ConnectionState.Open) { connection.Close(); }

            return result;
        }
        public static int Execute(string command)
        {
            var connection = new SqlConnection(Constr);
            var sqlCommand = new SqlCommand(command, connection);

            if (connection.State != ConnectionState.Open) { connection.Open(); }

            var result = sqlCommand.ExecuteNonQuery();

            if (connection.State == ConnectionState.Open) { connection.Close(); }

            return result;
        }
    }
}
