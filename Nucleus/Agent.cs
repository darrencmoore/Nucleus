using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Text;
using System.Net.Mail;


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
        /// <summary>
        /// Takes two lists.  
        /// One is the list of ContactId which are Guids (StartActvities). 
        /// The Guids are required to Post an Activity to Syspro
        /// The second is a list of sent bid Proposals (SentProposals)  that correspond to the entries in (StartActvities)
        /// I then build a xml string to post a activity to syspro.
        /// I also convert the report into binary for posting
        /// </summary>
        /// <param name="StartActvities"></param>
        /// <param name="SentProposals"></param>
        public void PostXmlForSyspro(List<Guid> StartActvities, List<Attachment> SentProposals)
        {
            // not needed for reporting. 
            // but needed for scanning
            //Encore.Utilities dll = new Encore.Utilities();
            //dll.Logon("ADMIN", " ", "TEST [TEST FOR 360 SHEET METAL LLC]", " ", Encore.Language.ENGLISH, 0, 0, " ");
            //Declaration
            DateTime bidSentDate = DateTime.Now;
            string formatDate = "yyyy-MM-dd";
            DateTime bidSentTime = DateTime.Now;
            string formatTime = "hh:mm:ss";
            int location = 0;
            byte[] propsalBinaryRep;

            try
            {
                foreach (Guid activity in StartActvities)
                {
                    StringBuilder activityXML = new StringBuilder();
                    activityXML.Append("Function RunMe(1,2)" + "& vbCrLf" +
                    activityXML.Append("Dim x As String" + "& vbCrLf" +

                    activityXML.Append("x = <PostActivity xmlns:xsd=\"http://www.w3.org/2001/XMLSchema-instance\" xsd:noNamespaceSchemaLocation=\"CMSTATDOC.XSD\">" + "& vbCrLf" +
                    activityXML.Append("x = x & <Item>" + "& vbCrLf" +
                    activityXML.Append("x = x & <ContactId>{" + activity + "}</ContactId>" + "& vbCrLf" +
                    activityXML.Append("x = x & <Activity>" + "& vbCrLf" +
                    activityXML.Append("x = x & <ActivityType>12</ActivityType>" + "& vbCrLf" +
                    activityXML.Append("x = x & <Private>N</Private>" + "& vbCrLf" +
                    activityXML.Append("x = x & <StartDate>" + bidSentDate.ToString(formatDate) + "</StartDate>" + "& vbCrLf" +
                    activityXML.Append("x = x & <StartTime>" + bidSentTime.ToString(formatTime) + "</StartTime>" + "& vbCrLf" +
                    activityXML.Append("x = x & <EndDate>" + bidSentDate.ToString(formatDate) + "</EndDate>" + "& vbCrLf" +
                    activityXML.Append("x = x & <EndTime>" + bidSentTime.ToString(formatTime) + "</EndTime>" + "& vbCrLf" +
                    activityXML.Append("x = x & <Subject>Bid Proposal</Subject>" + "& vbCrLf" +
                    //activityXML.Append("<Location>Head office</Location>");
                    //activityXML.Append("<Regarding>Sales call</Regarding>");
                    activityXML.Append("x = x & <Result>Email sent</Result>" + "& vbCrLf" +
                    activityXML.Append("x = x & <UserField1>User Field 1</UserField1>" + "& vbCrLf" +
                    activityXML.Append("x = x & <UserField2>User Field 2</UserField2>" + "& vbCrLf" +
                    activityXML.Append("x = x & <UserField3>User Field 3</UserField3>" + "& vbCrLf" +
                    //activityXML.Append("<Priority>9</Priority>");
                    //activityXML.Append("<FollowUpFlag>1</FollowUpFlag>");
                    //activityXML.Append("<FollowUpReqd>Y</FollowUpReqd>");
                    //activityXML.Append("<FollowUpDate>2008-01-03</FollowUpDate>");
                    //activityXML.Append("<FollowUpTime>08:30:00</FollowUpTime>");
                    //activityXML.Append("<AllDayEvent>N</AllDayEvent>");
                    //activityXML.Append("<ShowTimeAs>B</ShowTimeAs>");
                    //activityXML.Append("<TaskDueDate>2008-02-01</TaskDueDate>");
                    //activityXML.Append("<TaskPctComplete>50</TaskPctComplete>");
                    //activityXML.Append("<TaskStatus>4</TaskStatus>");
                    activityXML.Append("x = x & <Attachments>" + "& vbCrLf" +
                    activityXML.Append("x = x & <Attachment>" + "& vbCrLf")))))))))))))))))));
                    //foreach (Attachment attachment in SentProposals)
                    //{
                    //    activityXML.Append("x = x & <AttachmentName>" + SentProposals[location].Name + "</AttachmentName>" + "& vbCrLf" +
                    //    activityXML.Append("x = x & <AttachmentExt>pdf</AttachmentExt>" + "& vbCrLf"));
                    //    propsalBinaryRep = new byte[SentProposals[location].ContentStream.Length];
                    //    string xmlPropsalBinaryRepData;
                    //    SentProposals[location].ContentStream.Read(propsalBinaryRep, 0, (int)SentProposals[location].ContentStream.Length);
                    //    SentProposals[location].ContentStream.Close();
                    //    xmlPropsalBinaryRepData = System.Convert.ToBase64String(propsalBinaryRep, 0, propsalBinaryRep.Length);
                    //    activityXML.Append("x = x & <AttachmentData>" + xmlPropsalBinaryRepData + "</AttachmentData>" + "& vbCrLf");
                    //    location++;
                    //    break;
                    //}
                    activityXML.Append("x = x & </Attachment>" + "& vbCrLf" +
                    activityXML.Append("x = x & </Attachments>" + "& vbCrLf" +
                    activityXML.Append("x = x & <eSignature/>" + "& vbCrLf" +
                    activityXML.Append("x = x & </Activity>" + "& vbCrLf" +
                    activityXML.Append("x = x & </Item>" + "& vbCrLf" +
                    activityXML.Append("x = x & </PostActivity>"))))));

                    //ProLogicBid.Agent postActivity = new ProLogicBid.Agent();
                    //postActivity.SendEventDelegate("acivityPost", activityXML.ToString());




                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }            
        }
        #endregion
    }
}
