using System;
using System.Data;
using Oracle.ManagedDataAccess.Client;
using Oracle.ManagedDataAccess.Types;

namespace SeleniumFrameWork.Helpers
{
    public static class OLEDBHelpers
    {
        //Common Operation in database while it comes to automation testing
        //Read and Update,Delete Data by using System.Data namespace( Class Libraries)



        public static OracleConnection OracleConnection;
        public static OracleConnection DBConnect(string connectionString)
        { //try catch to handle any unknown exception
            try
            {
                OracleConnection = new OracleConnection(connectionString);
                OracleConnection.Open();
                return OracleConnection; //return the Oracle Connection which is Opened.


            }
            catch (Exception ex)
            {
                Console.Write("ERROR : :" + ex.Message);
            }
            return null; // If any exception have to return a null value

        }
        //Closing the Connection

        public static void DBclose(this OracleConnection OracleConnection)
        {
            try
            {     //calling Oracle Connection Close method to close the DB which is opened.
                OracleConnection.Close();

            }
            catch (Exception ex)
            {
                Console.Write("ERROR : :" + ex.Message);

            }
        }


        //Execution
        //DataTable Class to Return the data that is filled in the table
        public static DataTable ExecuteQuery(string queryString, string conn)
        {

            DBConnect(conn);
            DataSet dataset;
            try
            {
                //Checking the state of the connection
                if (OracleConnection == null || ((OracleConnection != null && (OracleConnection.State == ConnectionState.Closed || OracleConnection.State == ConnectionState.Broken))))
                    OracleConnection.Open();

                OracleDataAdapter dataAdaptor = new OracleDataAdapter();
                dataAdaptor.SelectCommand = new OracleCommand(queryString, OracleConnection);
                //Specifying the Command Type as text
                dataAdaptor.SelectCommand.CommandType = CommandType.Text;
                dataset = new DataSet();
                dataAdaptor.Fill(dataset, "table");
                OracleConnection.Close();
                return dataset.Tables["table"];

            }
            catch (Exception ex)
            {
                dataset = null;
                OracleConnection.Close();
                Logger.log("Error while Executing Query in OLE DB::" + ex.Message);
                return null;



            }
            finally
            {
                OracleConnection.Close();
                dataset = null;

            }
        }
    }
}
