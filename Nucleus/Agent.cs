using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Text;

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
                List<string> ContactList =  new List<string>(Agent_ContractContacts);
                //This is being sent the Bidding Report Application after post sorting
                List<string> _agentResponse = new List<string>();                
                foreach (string item in ContactList)
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
        
        /// <summary>
        /// Builds the Agent list to be sent to MainWindow.cs
        /// Checks to make sure that only one Account Name gets added with the LINQ statement
        /// </summary>
        /// <param name="item"></param>
        public void AgentContractContactsListItem(string item)
        {
            if(item.Contains("{ Header = Item Level 0 }") && !Agent_ContractContacts.Contains(item))
            {
                Agent_ContractContacts.Add(item);
            }
            if(item.Contains("{ Header = Item Level 1 }"))
            {
                Agent_ContractContacts.Add(item);
            }
            if (item.Contains("{ Header = Item Level 2 }"))
            {
                Agent_ContractContacts.Add(item);
            }                     
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
                _sqlCommand = new SqlCommand("usp_ContractHeader_pContract", _sqlConn);
                _sqlCommand.CommandType = CommandType.StoredProcedure;
                _sqlCommand.Parameters.Add(new SqlParameter("@Contract", contractId));
                sqlReader = _sqlCommand.ExecuteReader();
                               
                if (sqlReader.HasRows)
                {
                    while (sqlReader.Read())
                    {
                        AgentContractContactsListItem(sqlReader.GetString(1) + " " + sqlReader.GetString(2) + " " + "{ Header = Item Level 0 }"); //Account ID + Account Name
                        AgentContractContactsListItem(sqlReader.GetString(1) + " " + sqlReader.GetString(4) + " " + "{ Header = Item Level 1 }"); //Account ID + Contact Full Name
                        AgentContractContactsListItem(sqlReader.GetString(1) + "_" + sqlReader.GetGuid(3) + "_" + sqlReader.GetString(5) + " " + "{ Header = Item Level 2 }"); //Account ID + Contact Guid + Contact Email Address
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

        #region Post XML for Syspro
        public void PostXmlForSyspro(List<string> StartActvities)
        {
            // not needed for reporting. 
            // but needed for scanning
            //Encore.Utilities dll = new Encore.Utilities();
            //dll.Logon("ADMIN", " ", "TEST [TEST FOR 360 SHEET METAL LLC]", " ", Encore.Language.ENGLISH, 0, 0, " ");
            //Declaration
            //StringBuilder Document = new StringBuilder();

            //Building Document content
            //Document.Append("<?xml version=\"1.0\" encoding=\"Windows-1252\"?>");
            //Document.Append("<!-- Copyright 1994-2014 SYSPRO Ltd.-->");
            //Document.Append("<!--");
            //Document.Append("Sample XML for the Contact Activity Post Business Object");
            //Document.Append("-->");
            //Document.Append("<PostActivity xmlns:xsd=\"http://www.w3.org/2001/XMLSchema-instance\" xsd:noNamespaceSchemaLocation=\"CMSTATDOC.XSD\">");
            //Document.Append("<Item>");
            //Document.Append("<ContactId>{36303032-3031-3931-3135-313030373932}</ContactId>");
            //Document.Append("<Activity>");
            //Document.Append("<ActivityType>12</ActivityType>");
            //Document.Append("<Private>N</Private>");
            //Document.Append("<StartDate>2008-01-01</StartDate>");
            //Document.Append("<StartTime>07:00:00</StartTime>");
            //Document.Append("<EndDate>2008-01-01</EndDate>");
            //Document.Append("<EndTime>21:00:00</EndTime>");
            //Document.Append("<Subject>Update 00001</Subject>");
            //Document.Append("<Location>Head office</Location>");
            //Document.Append("<Regarding>Sales call</Regarding>");
            //Document.Append("<Result>Fax sent</Result>");
            //Document.Append("<UserField1>User Field 1</UserField1>");
            //Document.Append("<UserField2>User Field 2</UserField2>");
            //Document.Append("<UserField3>User Field 3</UserField3>");
            //Document.Append("<Priority>9</Priority>");
            //Document.Append("<FollowUpFlag>1</FollowUpFlag>");
            //Document.Append("<FollowUpReqd>Y</FollowUpReqd>");
            //Document.Append("<FollowUpDate>2008-01-03</FollowUpDate>");
            //Document.Append("<FollowUpTime>08:30:00</FollowUpTime>");
            //Document.Append("<AllDayEvent>N</AllDayEvent>");
            //Document.Append("<ShowTimeAs>B</ShowTimeAs>");
            //Document.Append("<TaskDueDate>2008-02-01</TaskDueDate>");
            //Document.Append("<TaskPctComplete>50</TaskPctComplete>");
            //Document.Append("<TaskStatus>4</TaskStatus>");
            //Document.Append("<Attendees>");
            //Document.Append("<Contact>james.paddington@bayside.com</Contact>");
            //Document.Append("<Contact>louise.granger@bayside.com</Contact>");
            //Document.Append("</Attendees>");
            //Document.Append("<Source>SOI</Source>");
            //Document.Append("<Body><![CDATA[{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fswiss\fcharset0 Arial;}}{\*\generator Msftedit 5.41.15.1507;}\viewkind4\uc1\pard\f0\fs20 Sales Fax Information\par}]]></Body>");
            //Document.Append("<Attachments>");
            //Document.Append("<Attachment>");
            //Document.Append("<AttachmentName>Specification.doc</AttachmentName>");
            //Document.Append("<AttachmentExt>doc</AttachmentExt>");
            //Document.Append("<AttachmentData><![CDATA[Binary content1]]></AttachmentData>");
            //Document.Append("</Attachment>");
            //Document.Append("</Attachments>");
            //Document.Append("<eSignature/>");
            //Document.Append("</Activity>");
            //Document.Append("</Item>");
            //Document.Append("</PostActivity>");


        }


        #endregion
    }
}
