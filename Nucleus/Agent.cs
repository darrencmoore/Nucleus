using System;
using System.IO;
using System.Net.Mime;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Linq;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Text;
using System.Net.Mail;
using Agent;
//using ProLogicReportingApplication;

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
        private string _operator;
        private string revisionNum;
        private string _availableOperator;
        string activity_AttachmentName;
        string activity_AttachmentExt;
        string activity_AttachmentData;
        private static string FUNCTIONALAREA_CMSP = "CMSP";


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
                if (item.Contains("BidSent = 0"))
                {
                    item = item.Replace("BidSent = 0", "False");
                }
                else if (item.Contains("BidSent = 1"))
                {                    
                    item = item.Replace("BidSent = 1", "True");
                }              
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
                        AgentContractContactsListItem(sqlReader.GetString(1) + " " + sqlReader.GetString(4) + " " + "{ Header = Item Level 1 }" + " " + "BidSent = " + sqlReader.GetInt32(15)); //Account ID + Contact Full Name + Bid Sent
                        AgentContractContactsListItem(sqlReader.GetString(1) + "_" + sqlReader.GetGuid(3) + "_" + sqlReader.GetString(5) + " " + "{ Header = Item Level 2 }" + " " + "{" + sqlReader.GetInt32(14) + "}"); //Account ID + Contact Guid + Contact Email Address + Contract Current Revision
                        // This is for debugging, if any fields get added 
                        //Console.WriteLine("{0}\t{1}\t{2}\t{3}\t{4}\t{5}\t{6}\t{7}\t{8}\t{9}\t{10}\t{11}\t{12}\t{13}\t{14}\t{15}", sqlReader.GetString(0), sqlReader.GetString(1), sqlReader.GetString(2), sqlReader.GetGuid(3), sqlReader.GetString(4), sqlReader.GetString(5), sqlReader.GetString(6), sqlReader.GetString(7), sqlReader.GetString(8), sqlReader.GetString(9), sqlReader.GetString(10), sqlReader.GetDateTime(11), sqlReader.GetDateTime(12), sqlReader.GetInt32(13), sqlReader.GetInt32(14), sqlReader.GetInt32(15));
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
        public void PostXmlForSyspro(List<Guid> StartActvities, List<Attachment> SentProposals, string _contractId, List<string> CurrentRevision)
        {            
            DateTime bidSentDate = DateTime.Now;
            string formatDate = "yyyy-MM-dd";
            DateTime bidSentTime = DateTime.Now;
            string formatTime = "hh:mm:ss";
            int location = 0;
            int currentRevisionLocation = 0;            
            FileStream inFile;
            FileStream fs;
            byte[] proposalBinaryRep;
            byte[] buffer;
            string contractId = _contractId;                      

            try
            {
                foreach (Guid activity in StartActvities)
                {                    
                    string activity_OpenPostActivityTag = "<PostActivity>";                    
                    string activity_OpenItemTag = "<Item>";
                    string activity_ContactGuid = "<ContactId>{" + activity + "}</ContactId>";
                    string activity_OpenActivityTag = "<Activity>";
                    string activity_ActivityType = "<ActivityType>12</ActivityType >";
                    string activity_PrivateTag = "<Private>N</Private>";
                    string activity_StartDate = "<StartDate>" + bidSentDate.ToString(formatDate) + "</StartDate>";
                    string activity_StartTime = "<StartTime>" + bidSentTime.ToString(formatTime) + "</StartTime>";
                    string activity_EndDate = "<EndDate>" + bidSentDate.ToString(formatDate) + "</EndDate>";
                    string activity_EndTime = "<EndTime>" + bidSentTime.ToString(formatTime) + "</EndTime>";
                    foreach(string currentRevisionNum in CurrentRevision)
                    {
                        revisionNum = currentRevisionNum;                        
                        currentRevisionLocation++;
                        break;
                    }
                    string activity_ActivitySubject = "<Subject>" + "Contract:" + contractId + " " + "Revision:" + revisionNum + "</Subject>";
                    string activity_ActivityResult = "<Result>Email sent</Result>";
                    string activity_UserField1 = "<UserField1>User Field 1</UserField1>";
                    string activity_UserField2 = "<UserField1>User Field 2</UserField1>";
                    string activity_UserField3 = "<UserField1>User Field 3</UserField1>";
                    string activity_Source = "<Source>DotNet</Source>";
                    string activity_OpenAttachmentsTag = "<Attachments>";
                    string activity_OpenAttachmentTag = "<Attachment>";
                    //foreach (Attachment attachment in SentProposals)
                    //{
                        //string _attachmentName = "<AttachmentName>" + SentProposals[location].Name + "</AttachmentName>";
                        //activity_AttachmentName = _attachmentName;
                        //string _attachmentExt = "<AttachmentExt>pdf</AttachmentExt>";
                        //activity_AttachmentExt = _attachmentExt;
                        //proposalBinaryRep = File.ReadAllBytes(SentProposals[location].Name);
                        //using (fs = new FileStream(SentProposals[location].Name, FileMode.Open, FileAccess.Read))
                        //{
                        //    buffer = new byte[fs.Length];
                        //    fs.Read(buffer, 0, (int)fs.Length);
                        //}
                            //long bytesRead = proposalBinaryRep
                            //inFile = SentProposals[location].ContentStream;
                            //var path = Path.GetFullPath(SentProposals[location]);
                            //inFile = new FileStream(path, FileMode.Open, FileAccess.Read);
                            //proposalBinaryRep = new Byte[inFile.Length];
                            //proposalBinaryRep = File.ReadAllBytes(attachment);
                            //proposalBinaryRep = SentProposals[location].ContentStream.ReadByte();

                            //var binaryString = ToBinary(ConvertToByte(SentProposals[location].ContentStream, Encoding.UTF32);
                            //UInt32 xmlProposalBinaryRepData;
                            //SentProposals[location].ContentStream.Read(proposalBinaryRep, 0, (int)SentProposals[location].ContentStream.Length);
                            //SentProposals[location].ContentStream.Close();
                            //xmlProposalBinaryRepData = System.Convert.ToUInt32(proposalBinaryRep);//, 0, proposalBinaryRep.Length); 
                            //string _attachmentData = "<AttachmentData>" + SentProposals[location].Name + "</AttachmentData>";
                        //activity_AttachmentData = _attachmentData;
                        //location++;
                        //break;                       
                    //}
                    string activity_ClosingAttachmentTag = "</Attachment>";
                    string activity_ClosingAttachmentsTag = "</Attachments>";
                    string activity_ESignature = "<eSignature/>";
                    string activity_ClosingActivityTag = "</Activity>";
                    string activity_ClosingItemTag = "</Item>";
                    string activity_PostTag = "</PostActivity>";

                    StringBuilder activity_XmlParam = new StringBuilder();
                    activity_XmlParam.Append("<PostActivity>");
                    activity_XmlParam.Append("<Parameters>");
                    activity_XmlParam.Append("<ActionType>A</ActionType>");
                    activity_XmlParam.Append("<AttendeeIdType>{email}</AttendeeIdType>");
                    activity_XmlParam.Append("<ApplyIfEntireDocumentValid>N</ApplyIfEntireDocumentValid>");
                    activity_XmlParam.Append("<IgnoreAttachmentsOnChange>N</IgnoreAttachmentsOnChange>");
                    activity_XmlParam.Append("</Parameters>");
                    activity_XmlParam.Append("</PostActivity>");

                    StringBuilder activity_XmlDoc = new StringBuilder();                    
                    activity_XmlDoc.Append(activity_OpenPostActivityTag);
                    activity_XmlDoc.Append(activity_OpenItemTag);
                    activity_XmlDoc.Append(activity_ContactGuid);
                    activity_XmlDoc.Append(activity_OpenActivityTag);
                    activity_XmlDoc.Append(activity_ActivityType);
                    activity_XmlDoc.Append(activity_PrivateTag);
                    activity_XmlDoc.Append(activity_StartDate);
                    activity_XmlDoc.Append(activity_StartTime);
                    activity_XmlDoc.Append(activity_EndDate);
                    activity_XmlDoc.Append(activity_EndTime);
                    activity_XmlDoc.Append(activity_ActivitySubject);
                    //activity_XmlDoc.Append("XMLDoc =" + "<Location>Head office</Location>");
                    //activity_XmlDoc.Append("XMLDoc =" + "<Regarding>Sales call</Regarding>");
                    activity_XmlDoc.Append(activity_ActivityResult);
                    activity_XmlDoc.Append(activity_UserField1);
                    activity_XmlDoc.Append(activity_UserField2);
                    activity_XmlDoc.Append(activity_UserField3);
                    activity_XmlDoc.Append(activity_Source);
                    //activity_XmlDoc.Append("XMLDoc =" + "<Priority>9</Priority>");
                    //activity_XmlDoc.Append("XMLDoc =" + "<FollowUpFlag>1</FollowUpFlag>");
                    //activity_XmlDoc.Append("XMLDoc =" + "<FollowUpReqd>Y</FollowUpReqd>");
                    //activity_XmlDoc.Append("XMLDoc =" + "<FollowUpDate>2008-01-03</FollowUpDate>");
                    //activity_XmlDoc.Append("XMLDoc =" + "<FollowUpTime>08:30:00</FollowUpTime>");
                    //activity_XmlDoc.Append("XMLDoc =" + "<AllDayEvent>N</AllDayEvent>");
                    //activity_XmlDoc.Append("XMLDoc =" + "<ShowTimeAs>B</ShowTimeAs>");
                    //activity_XmlDoc.Append("XMLDoc =" + "<TaskDueDate>2008-02-01</TaskDueDate>");
                    //activity_XmlDoc.Append("XMLDoc =" + "<TaskPctComplete>50</TaskPctComplete>");
                    //activity_XmlDoc.Append("XMLDoc =" + "<TaskStatus>4</TaskStatus>");
                    activity_XmlDoc.Append(activity_OpenAttachmentsTag);
                    activity_XmlDoc.Append(activity_OpenAttachmentTag);
                    activity_XmlDoc.Append(activity_AttachmentName);
                    activity_XmlDoc.Append(activity_AttachmentExt);
                    activity_XmlDoc.Append(activity_AttachmentData);
                    activity_XmlDoc.Append(activity_ClosingAttachmentTag);
                    activity_XmlDoc.Append(activity_ClosingAttachmentsTag);
                    activity_XmlDoc.Append(activity_ESignature);
                    activity_XmlDoc.Append(activity_ClosingActivityTag);
                    activity_XmlDoc.Append(activity_ClosingItemTag);
                    activity_XmlDoc.Append(activity_PostTag);

                    _availableOperator = GetOperator();

                    Encore.Utilities sessionInstance = new Encore.Utilities();
                    string sessionID = sessionInstance.Logon(_availableOperator, "uP1ndkE9", "TEST", " ", Encore.Language.ENGLISH, 0, 0, " ");
                    Encore.Transaction postActivity = new Encore.Transaction();
                    postActivity.Post(sessionID, "CMSTAT", activity_XmlParam.ToString(), activity_XmlDoc.ToString());
                    sessionInstance.Logoff(sessionID);

                    ClearOperator(_availableOperator);                                       
                }                
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());              
            }            
        }
        #endregion

        #region Get Available Operator
        private string GetOperator()
        {
            try
            {
                SqlDataReader sqlReader;
                _connStr = ConfigurationManager.ConnectionStrings["SYSPRO_SQL_SERVER"].ConnectionString;
                DBOpenConnection();
                _sqlCommand = new SqlCommand("usp_ENetUserGet", _sqlConn);
                _sqlCommand.CommandType = CommandType.StoredProcedure;
                _sqlCommand.Parameters.Add(new SqlParameter("@FunctionalArea", FUNCTIONALAREA_CMSP));                
                sqlReader = _sqlCommand.ExecuteReader();

                if (sqlReader.HasRows)
                {
                    while (sqlReader.Read())
                    {
                        _operator = sqlReader.GetString(0);
                    }
                }
                
                return _operator;
            }
            catch (Exception exc)
            {
                MessageBox.Show("All ENet Users currently in use.  Please contact your Administrator.");
            }

            return null;
        }
        #endregion

        #region 
        private void ClearOperator(string usedOperator)
        {
            SqlDataReader sqlReader;
            _connStr = ConfigurationManager.ConnectionStrings["SYSPRO_SQL_SERVER"].ConnectionString;
            DBOpenConnection();
            _sqlCommand = new SqlCommand("usp_ENetUserClear", _sqlConn);
            _sqlCommand.CommandType = CommandType.StoredProcedure;
            _sqlCommand.Parameters.Add(new SqlParameter("@UserID", usedOperator));            
            sqlReader = _sqlCommand.ExecuteReader();
        }

        #endregion
    }
}
