﻿using System;
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

        public List<string> Agent_ContractContacts = new List<string>();
        public List<string> Agent_ContractContactsResponse
        {
            get
            {
                // hash set eliminiates duplicate Account names for 
                //the _proLogic_zContractContacts list based on contract
                var ContactHash = new HashSet<string>(Agent_ContractContacts);
                //THis is being sent the Bidding Report Application after post sorting the hash set
                List<string> _agentResponse = new List<string>();                
                foreach (string item in ContactHash)
                {                  
                    if (item.Contains("{ Header = Item Level 0 }")) // parent
                    {
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
        
        public void AgentContractContactsListItem(string item)
        {
            Agent_ContractContacts.Add(item);
        }
        #endregion

        #region Database Connections and SELECT Statement for tree view
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
        /// Runs the usp_ContractHeader_pContact stored procedure 
        /// </summary>
        /// <param name="sql"></param>
        /// <returns> Returns a list for the treeview </returns>
        public SqlDataReader GetContacts(string contractId)
        {
            try
            {
                SqlDataReader sqlReader;                                
                _connStr = ConfigurationManager.ConnectionStrings["SYSPRO_SQL_SERVER"].ConnectionString;                                
                DBOpenConnection();
                _sqlCommand = new SqlCommand("usp_ContractHeader_pContact", _sqlConn);
                _sqlCommand.CommandType = CommandType.StoredProcedure;
                _sqlCommand.Parameters.Add(new SqlParameter("@Contract", contractId));
                sqlReader = _sqlCommand.ExecuteReader();
                               
                if (sqlReader.HasRows)
                {
                    while (sqlReader.Read())
                    {
                        AgentContractContactsListItem(sqlReader.GetString(1) + " " + sqlReader.GetString(2) + " " + "{ Header = Item Level 0 }"); //Account ID + Account Name
                        AgentContractContactsListItem(sqlReader.GetString(1) + " " + sqlReader.GetString(4) + " " + "{ Header = Item Level 1 }"); //Account ID + Contact Full Name
                        AgentContractContactsListItem(sqlReader.GetString(5) + " " + "{ Header = Item Level 2 }"); //Contact Email Address "{ Header = Item Level 2 }"
                        // This is for debugging, if any fields get added 
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

        #region Select Statement for report preview
        /// <summary>
        /// Runs the usp_ContractDetail_pContract_Account stored procedure 
        /// </summary>
        /// <param name="contractId"></param>
        /// <param name="accountId"></param>
        /// <returns>Returns a datatable for the Crystal Report report source</returns>
        public DataTable ReportPreview(string contractId, string accountId)
        {
            try
            {
                DataTable reportPreview = new DataTable();
                _connStr = ConfigurationManager.ConnectionStrings["SYSPRO_SQL_SERVER"].ConnectionString;
                DBOpenConnection();
                _sqlCommand = new SqlCommand("usp_ContractDetail_pContract_Account", _sqlConn);
                _sqlCommand.Parameters.Add(new SqlParameter("@Contract", contractId));
                _sqlCommand.Parameters.Add(new SqlParameter("@Account", accountId));
                _sqlCommand.CommandType = CommandType.StoredProcedure;
                using (SqlDataAdapter adapReportPreview = new SqlDataAdapter(_sqlCommand))
                {
                    adapReportPreview.Fill(reportPreview);
                }
                return reportPreview;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            return null;
        }
        #endregion
    }
}
