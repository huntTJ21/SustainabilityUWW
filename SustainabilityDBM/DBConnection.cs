using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SustainabilityDBM
{
    class DBConnection
    {
        // Member Fields
        private bool __IsConnected = false;

        // Getters and Setters
        public string Server { get; set; }
        public string DBName { get; set; }
        public SqlConnection Connection { get; set; }

        // Member Functions
        private string BuildConnectionString()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat("Server={0}; Database= {1}; Integrated Security=SSPI;", Server, DBName);
            return sb.ToString();
        }

        public bool IsConnected()
        {
            return __IsConnected;
        }

        public bool Connect()
        {
            if (Connection == null)
            {
                Connection = new SqlConnection(BuildConnectionString());
            }

            try
            {
                Connection.Open();
                __IsConnected = true;
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public bool Disconnect()
        {
            try
            {
                Connection.Close();
                __IsConnected = false;
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        // Object Instantiation
        public DBConnection(String server, String dbName, bool autoConnect = false)
        {
            this.Server = server;
            this.DBName = dbName;
            this.Connection = new SqlConnection(BuildConnectionString());
            if (autoConnect)
            {
                Connect();
            }
        }
    }
}
