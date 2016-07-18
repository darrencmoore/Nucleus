using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Diagnostics;
using System.Windows;


/// <summary>
/// Created On: 6/22/2016
/// By: Darren Moore
/// </summary>

namespace Nucleus
{
    public class Agent
    {
        private SqlConnection _sqlConn;
        private SqlCommand _sqlCommand;
        private string _connStr;
                

#region Getter/Setter ZContactContractsSelect

        public List<string> _proLogic_zContractContacts = new List<string>();
        public List<string> getProLogic_zContractContacts
        {
            get
            {
                // hash set eliminiates duplicate Account names for 
                //the _proLogic_zContractContacts list based on contract
                var ContactHash = new HashSet<string>(_proLogic_zContractContacts);
                //THis is being sent the Bidding Report Application after post sorting the hash set
                List<string> _agentResponse = new List<string>();                
                foreach (string item in ContactHash)
                {  
                    if (item.Contains("{ Header = Item Level 0 }")) // parent
                    {
                        //item.Replace("{ Header = Item Level 0 }", "");
                        _agentResponse.Add(item);
                        continue;
                    }
                    if(item.Contains("{ Header = Item Level 1 }") || item.Contains("{ Header = Item Level 2 }")) // child
                    {
                        _agentResponse.Add(item);
                    }                   
                }               

                return _agentResponse;
            }
        }        
        
        public void AddzContractContactsLineItem(string item)
        {
            _proLogic_zContractContacts.Add(item);
        }
#endregion

#region Database Connections and SELECT, INSERT, DELETE      

        /// <summary>
        /// Opens Database Connection
        /// </summary>
        /// <returns></returns>
        public bool DBOpenConnection()
        {
            try
            {                
                _sqlConn = new SqlConnection(_connStr);
                _sqlConn.Open();
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        /// <summary>
        /// Closes Database Connection
        /// </summary>
        /// <returns></returns>
        public bool DBCloseConnection()
        {
            try
            {
                _sqlConn = new SqlConnection(_connStr);
                if (_sqlConn.State != ConnectionState.Closed)
                    _sqlConn.Close();
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        /// <summary>
        /// Takes any select statement passed to it. 
        /// </summary>
        /// <param name="sql"></param>
        /// <returns></returns>
        public SqlDataReader Select(string ID)
        {
            try
            {
                SqlDataReader sqlReader;
                                
                _connStr = ConfigurationManager.ConnectionStrings["SYSPRO_SQL_SERVER"].ConnectionString;                                
                DBOpenConnection();
                _sqlCommand = new SqlCommand("usp_ContractHeader_ZccPcmPpcmcCeZcbc", _sqlConn); //usp_ProjectHeader_ZccPcmPpcmcZcbc
                _sqlCommand.CommandType = CommandType.StoredProcedure;
                _sqlCommand.Parameters.Add(new SqlParameter("@Contract", ID));
                sqlReader = _sqlCommand.ExecuteReader();
                               
                if (sqlReader.HasRows)
                {
                    while (sqlReader.Read())
                    {
                        AddzContractContactsLineItem(sqlReader.GetString(2) + " " + "{ Header = Item Level 0 }"); //Account Name 
                        AddzContractContactsLineItem(sqlReader.GetString(3) + " " + "{ Header = Item Level 1 }"); //Contact Full Name 
                        AddzContractContactsLineItem(sqlReader.GetString(4) + " " + "{ Header = Item Level 2 }"); //Contact Email Address "{ Header = Item Level 2 }"
                        //Console.WriteLine("{0}\t{1}\t{2}\t{3}\t{4}\t{5}\t{6}\t{7},\t{8},\t{9},\t{10}", sqlReader.GetString(0), sqlReader.GetGuid(1), sqlReader.GetString(2), sqlReader.GetString(3), sqlReader.GetString(4), sqlReader.GetString(5), sqlReader.GetString(6), sqlReader.GetString(7), sqlReader.GetString(8), sqlReader.GetDateTime(9), sqlReader.GetDateTime(10));
                    }
                }
                else
                {
                    Console.WriteLine("No Rows Found");
                }

                return sqlReader;
            }
            catch (Exception ex)
            {
                return null;
            }
        }
#endregion
    }
}
