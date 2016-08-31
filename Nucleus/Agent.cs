using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Linq;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Text;
using System.Net.Mail;
using Agent;




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
        string activity_AttachmentName;
        string activity_AttachmentExt;
        string activity_AttachmentData;

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
                    string activity_DimX = "Dim x As String = ' '";
                    string activity_FunctionName = "Function ActivityPost(a,b)";
                    string activity_DimXmlOut = "Dim XMLOut";
                    string activity_DimXmlParam = "Dim XMLParam";
                    string activity_DimXmlDoc = "Dim XMLDoc";
                    string activity_OpenPostActivityTag = "<PostActivity>";                    
                    string activity_OpenItemTag = "<Item>";
                    string activity_ContactGuid = "<ContactId>{" + activity + "}</ContactId>";
                    string activity_OpenActivityTag = "<Activity>";
                    string activity_ActivityType = "< ActivityType > 12 </ ActivityType >";
                    string activity_PrivateTag = "<Private>N</Private>";
                    string activity_StartDate = "<StartDate>" + bidSentDate.ToString(formatDate) + "</StartDate>";
                    string activity_StartTime = "<StartTime>" + bidSentTime.ToString(formatTime) + "</StartTime>";
                    string activity_EndDate = "<EndDate>" + bidSentDate.ToString(formatDate) + "</EndDate>";
                    string activity_EndTime = "<EndTime>" + bidSentTime.ToString(formatTime) + "</EndTime>";
                    string activity_AcivitySubject = "<Subject>Bid Proposal</Subject>";
                    string activity_ActivityResult = "<Result>Email sent</Result>";
                    string activity_UserField1 = "<UserField1>User Field 1</UserField1>";
                    string activity_UserField2 = "<UserField1>User Field 2</UserField1>";
                    string activity_UserField3 = "<UserField1>User Field 3</UserField1>";
                    string activity_Source = "<Source>DotNet</Source>";
                    string activity_OpenAttachmentsTag = "<Attachments>";
                    string activity_OpenAttachmentTag = "<Attachment>";
                    foreach (Attachment attachment in SentProposals)
                    {
                        string _attachmentName = "<AttachmentName>" + SentProposals[location].Name + "</AttachmentName>";
                        activity_AttachmentName = _attachmentName;
                        string _attachmentExt = "<AttachmentExt>pdf</AttachmentExt>";
                        activity_AttachmentExt = _attachmentExt;
                        propsalBinaryRep = new byte[SentProposals[location].ContentStream.Length];
                        string xmlPropsalBinaryRepData;
                        SentProposals[location].ContentStream.Read(propsalBinaryRep, 0, (int)SentProposals[location].ContentStream.Length);
                        SentProposals[location].ContentStream.Close();
                        xmlPropsalBinaryRepData = System.Convert.ToBase64String(propsalBinaryRep, 0, propsalBinaryRep.Length);
                        string _attachmentData = "<AttachmentData>" + xmlPropsalBinaryRepData + "</AttachmentData>";
                        activity_AttachmentData = _attachmentData;
                        location++;
                        break;                       
                    }
                    string activity_ClosingAttachmentTag = "</Attachment>";
                    string activity_ClosingAttachmentsTag = "</Attachments>";
                    string activity_ESignature = "<eSignature/>";
                    string activity_ClosingActivityTag = "</Activity>";
                    string activity_ClosingItemTag = "</Item>";
                    string activity_PostTag = "</PostActivity>";

                    StringBuilder activity_XmlParam = new StringBuilder();
                    activity_XmlParam.Append(activity_DimX);
                    activity_XmlParam.Append(activity_FunctionName);
                    activity_XmlParam.Append(activity_DimXmlOut);
                    activity_XmlParam.Append(activity_DimXmlParam);
                    activity_XmlParam.Append(activity_DimXmlDoc);
                    activity_XmlParam.Append("XMLParam =" + "<PostActivity>");
                    activity_XmlParam.Append("XMLParam =" + "<Parameters>");
                    activity_XmlParam.Append("XMLParam =" + "<ActionType>A</ActionType>");
                    activity_XmlParam.Append("XMLParam =" + "<AttendeeIdType>{email}</AttendeeIdType>");
                    activity_XmlParam.Append("XMLParam =" + "<ApplyIfEntireDocumentValid>N</ApplyIfEntireDocumentValid>");
                    activity_XmlParam.Append("XMLParam =" + "<IgnoreAttachmentsOnChange>N</IgnoreAttachmentsOnChange>");
                    activity_XmlParam.Append("XMLParam =" + "</Parameters>");
                    activity_XmlParam.Append("XMLParam =" + "</PostActivity>");

                    StringBuilder activity_XmlDoc = new StringBuilder();                    
                    activity_XmlDoc.Append("XMLDoc =" + activity_OpenPostActivityTag);
                    activity_XmlDoc.Append("XMLDoc =" + activity_OpenItemTag);
                    activity_XmlDoc.Append("XMLDoc =" + activity_ContactGuid);
                    activity_XmlDoc.Append("XMLDoc =" + activity_OpenActivityTag);
                    activity_XmlDoc.Append("XMLDoc =" + activity_ActivityType);
                    activity_XmlDoc.Append("XMLDoc =" + activity_PrivateTag);
                    activity_XmlDoc.Append("XMLDoc =" + activity_StartDate);
                    activity_XmlDoc.Append("XMLDoc =" + activity_StartTime);
                    activity_XmlDoc.Append("XMLDoc =" + activity_EndDate);
                    activity_XmlDoc.Append("XMLDoc =" + activity_EndTime);
                    activity_XmlDoc.Append("XMLDoc =" + activity_AcivitySubject);
                    //activity_XmlDoc.Append("XMLDoc =" + "<Location>Head office</Location>");
                    //activity_XmlDoc.Append("XMLDoc =" + "<Regarding>Sales call</Regarding>");
                    activity_XmlDoc.Append("XMLDoc =" + activity_ActivityResult);
                    activity_XmlDoc.Append("XMLDoc =" + activity_UserField1);
                    activity_XmlDoc.Append("XMLDoc =" + activity_UserField2);
                    activity_XmlDoc.Append("XMLDoc =" + activity_UserField3);
                    activity_XmlDoc.Append("XMLDoc =" + activity_Source);
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
                    activity_XmlDoc.Append("XMLDoc =" + activity_OpenAttachmentsTag);
                    activity_XmlDoc.Append("XMLDoc =" + activity_OpenAttachmentTag);
                    activity_XmlDoc.Append("XMLDoc =" + activity_AttachmentName);
                    activity_XmlDoc.Append("XMLDoc =" + activity_AttachmentExt);
                    activity_XmlDoc.Append("XMLDoc =" + activity_AttachmentData);
                    activity_XmlDoc.Append("XMLDoc =" + activity_ClosingAttachmentTag);
                    activity_XmlDoc.Append("XMLDoc =" + activity_ClosingAttachmentsTag);
                    activity_XmlDoc.Append("XMLDoc =" + activity_ESignature);
                    activity_XmlDoc.Append("XMLDoc =" + activity_ClosingActivityTag);
                    activity_XmlDoc.Append("XMLDoc =" + activity_ClosingItemTag);
                    activity_XmlDoc.Append("XMLDoc =" + activity_PostTag);

                    //Lay the structure for getting a new Operator
                     
                    Encore.Utilities sessionInstance = new Encore.Utilities();
                    string sessionID = sessionInstance.Logon("ENET_CMSP01", "uP1ndkE9", "TEST", " ", Encore.Language.ENGLISH, 0, 0, " ");                    
                    Encore.Transaction postActivity = new Encore.Transaction();
                    postActivity.Post(sessionID, "CMSTAT", activity_XmlParam.ToString(), activity_XmlDoc.ToString());
                    sessionInstance.Logoff(sessionID);                    
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
