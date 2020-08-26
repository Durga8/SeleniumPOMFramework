using SeleniumFrameWork.Config;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

//Initial Comments
namespace SeleniumFrameWork.Helpers
{
    public static class DBHelpers
    {
        public static SqlConnection sqlConnection;
        //Open the connection
        public static SqlConnection DBConnect(string Conn)
        {
            try
            {

                sqlConnection = new SqlConnection(Conn);
                sqlConnection.Open();
                return sqlConnection;
            }

            catch (Exception e)
            {
                Logger.log("ERROR::" + e.Message);

            }
            return null;
        }
        public static void DBClose()
        {
            try
            {
                sqlConnection.Close();


            }
            catch (Exception e)
            {
                Logger.log("ERROR::" + e.Message);
            }

        }
        //Execution

        public static DataTable ExecuteQuery(string queryString, string conn)
        {
            // string Conn = GlobalResource.connectionString;
            //string Conn = ExcelHelpers.getParameter("connectionString");
            DBConnect(conn);
            DataSet dataset;
            try
            {  //Checking the state of the connection
                if (sqlConnection == null || ((sqlConnection != null && (sqlConnection.State == ConnectionState.Closed ||
                  sqlConnection.State == ConnectionState.Broken))))
                    sqlConnection.Open();

                SqlDataAdapter Adaptor = new SqlDataAdapter();
                Adaptor.SelectCommand = new SqlCommand(queryString, sqlConnection);
                Adaptor.SelectCommand.CommandType = CommandType.Text;

                dataset = new DataSet();
                Adaptor.Fill(dataset, "table");
                sqlConnection.Close();
                return dataset.Tables["table"];




            }

            catch (Exception e)
            {
                dataset = null;
                sqlConnection.Close();
                Logger.log("ERROR::" + e.Message);
                return null;

            }

            finally
            {
                sqlConnection.Close();
                dataset = null;


            }
        }

        //Read Data from DB
        



    }
}
