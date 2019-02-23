﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data.Odbc;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using JUST.Shared.DatabaseRepository;
using JUST.Shared.Classes;

namespace JUST_PONotifier
{
    public class MainClass
    {
        /* version 1.03  */
        private const string debug = "debug";
        private const string live = "live";
        private const string monitor = "monitor";

        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private static string Uid;
        private static string Pwd;
        private static string FromEmailAddress;
        private static string FromEmailPassword;
        private static string FromEmailSMTP;
        private static int? FromEmailPort;
        private static string Mode;
        private static string[] MonitorEmailAddresses;
        private static ArrayList ValidModes = new ArrayList() { debug, live, monitor };
        private static string POAttachmentBasePath;

        private static string EmailSubject = "Purchase Order {0} Received from {1}";
        private static string MessageBodyFormat = @"
    <body style = ""margin-left: 20px; margin-right:20px"" >
        <hr/>
        <h2> Purchase Order {0} Received</h2>
        <hr/>
        <p>
        Received By: {1}<br/>
        Received Date: {2}<br/>
        Located in bin: {3}<br />
        Job #{4}   {5}<br/>
        Workorder: {6}   {7}<br/>
        Customer: {8}<br/>
        Vendor Name: {9}<br/>
        Buyer: {10}<br/>
        Notes: {11}<br/>

        <table style = ""width:50%; text-align: left"" border=""1"" cellpadding=""10"" cellspacing=""0"">
            <tr style = ""background-color: cyan"" >
                <th>Item Number</th>
                <th>Description</th>
                <th style = ""text-align: center"">Qty</th>
            </tr>";

        private static string messageBodyTableItem = @"<tr>
                <td>{0}</td>
                <td>{1}</td>
                <td style=""text-align: center"">{2}</td>
            </tr>";
        private static string messageBodyTail = @"</table></body>";

        static void Main(string[] args)
        {
            try
            {
                log.Info("[Main] Starting up at " + DateTime.Now);

                getConfiguration();

                ProcessPOData();

                log.Info("[Main] Completion at " + DateTime.Now);
            }
            catch (Exception ex)
            {
                log.Error("[Main] Error: " + ex.Message);
            }
        }

        private static void getConfiguration()
        {
            Uid = ConfigurationManager.AppSettings["Uid"];
            Pwd = ConfigurationManager.AppSettings["Pwd"];
            FromEmailAddress = ConfigurationManager.AppSettings["FromEmailAddress"];
            FromEmailPassword = ConfigurationManager.AppSettings["FromEmailPassword"];
            FromEmailSMTP = ConfigurationManager.AppSettings["FromEmailSMTP"];
            FromEmailPort = Convert.ToInt16(ConfigurationManager.AppSettings["FromEmailPort"]);

            Mode = ConfigurationManager.AppSettings["Mode"].ToLower();
            var MonitorEmailAddressList = ConfigurationManager.AppSettings["MonitorEmailAddress"];
            if (MonitorEmailAddressList.Length > 0)
            {
                char[] delimiterChars = { ';', ',' };
                MonitorEmailAddresses = MonitorEmailAddressList.Split(delimiterChars);
            }

            POAttachmentBasePath = ConfigurationManager.AppSettings["POAttachmentBasePath"];

            #region Validate Configuration Data
            var errorMessage = new StringBuilder();
            if (String.IsNullOrEmpty(Uid))
            {
                errorMessage.Append("User ID (Uid) is Required");
            }

            if (String.IsNullOrEmpty(Pwd))
            {
                errorMessage.Append("Password (Pwd) is Required");
            }

            if (String.IsNullOrEmpty(FromEmailAddress))
            {
                errorMessage.Append("From Email Address (FromEmailAddress) is Required");
            }

            if (String.IsNullOrEmpty(FromEmailPassword))
            {
                errorMessage.Append("From Email Password (FromEmailPassword) is Required");
            }

            if (String.IsNullOrEmpty(FromEmailSMTP))
            {
                errorMessage.Append("From Email SMTP (FromEmailSMTP) address is Required");
            }

            if (!FromEmailPort.HasValue)
            {
                errorMessage.Append("From Email Port (FromEmailPort) is Required");
            }

            if (String.IsNullOrEmpty(Mode))
            {
                errorMessage.Append("Mode is Required");
            }

            if (!ValidModes.Contains(Mode.ToLower()))
            {
                errorMessage.Append(String.Format("{0} is not a valid Mode.  Valid modes are 'debug', 'live' and 'monitor'", Mode));
            }

            if (Mode == monitor)
            {
                log.Info("checking MontorEmailAddressList");
                if (MonitorEmailAddresses == null || MonitorEmailAddresses.Length == 0)
                {
                    errorMessage.Append("Monitor Email Address is Required in monitor mode");
                }
                log.Info("finished checking MontorEmailAddressList");
            }

            if (String.IsNullOrEmpty(POAttachmentBasePath))
            {
                errorMessage.Append("Root Path to Attachments (AttachmentBasePath) is Required");
            }

            if (errorMessage.Length > 0)
            {
                throw new Exception(errorMessage.ToString());
            }
            #endregion

        }

        private static void ProcessPOData()
        {
            try
            {
                OdbcConnection cn;
//                OdbcCommand cmd;
//                string POQuery;
                var notifiedlist = new ArrayList();

                // user_1 = receiving rack location
                // user_2 = Receiver
                // user_3 = Received Date
                // user_4 = Bin Cleared Date
                // user_5 = Notified
                //                POQuery = "Select icpo.buyer, icpo.ponum, icpo.user_1, icpo.user_2, icpo.user_3, icpo.user_4, icpo.user_5, icpo.defaultjobnum, vendor.name as vendorName, icpo.user_6, icpo.defaultworkorder, icpo.attachid from icpo inner join vendor on vendor.vennum = icpo.vennum where icpo.user_3 is not null and icpo.user_5 = 0 order by icpo.ponum asc";

                OdbcConnectionStringBuilder just = new OdbcConnectionStringBuilder
                {
                    Driver = "ComputerEase"
                };
                just.Add("Dsn", "Company 0");
                just.Add("Uid", Uid);
                just.Add("Pwd", Pwd);

                cn = new OdbcConnection(just.ConnectionString);
//                cmd = new OdbcCommand(POQuery, cn);
                cn.Open();
                log.Info("[ProcessPOData] Connection to database opened successfully");
                var dbRepository = new DatabaseRepository(cn, log, POAttachmentBasePath);
                
                List<PurchaseOrder> purchaseOrdersToNotify;
                try
                {
                    purchaseOrdersToNotify = dbRepository.GetPurchaseOrdersToNotify();
                    log.Info("purchaseOrdersToNotify found " + purchaseOrdersToNotify.Count.ToString() + " items.");

                    foreach (PurchaseOrder po in purchaseOrdersToNotify)
                    {
                        var job = dbRepository.GetEmailBodyInformation(po.JobNumber, po.PurchaseOrderNumber, po.WorkOrderNumber);
                        var buyerEmployee = GetEmployeeInformation(dbRepository.GetEmployees(), po.Buyer);
                        var projectManagerEmployee = job.ProjectManagerName.Length > 0 ? GetEmployeeInformation(dbRepository.GetEmployees(), job.ProjectManagerName) : new Employee();

                        log.Info("[ProcessPOData] ----------------- Found PO Number " + po.PurchaseOrderNumber + " -------------------");

                        var emailSubject = String.Format(EmailSubject, po.PurchaseOrderNumber, po.Vendor);
                        var emailBody = FormatEmailBody(po.ReceivedOnDate, po.PurchaseOrderNumber, po.ReceivedBy, po.Bin, buyerEmployee.Name, po.Vendor, job, po.Notes);

                        ArrayList primaryRecipients = new ArrayList();
                        ArrayList bccList = new ArrayList();
                        if ((Mode == live) || (Mode == monitor))
                        {
                            primaryRecipients.Add(buyerEmployee.EmailAddress);
                            if (projectManagerEmployee.EmailAddress.Length > 0 && buyerEmployee.EmailAddress != projectManagerEmployee.EmailAddress)
                            {
                                primaryRecipients.Add(projectManagerEmployee.EmailAddress);
                            }
                        }

                        if (((Mode == monitor) || (Mode == debug)) &&
                            (MonitorEmailAddresses != null && MonitorEmailAddresses.Length > 0))
                        {
                            foreach(string monitorEmailAddress in MonitorEmailAddresses)
                            {
                                bccList.Add(monitorEmailAddress);
                            }
                        }

                        if ((primaryRecipients.Count == 0) && (bccList.Count > 0))
                        {
                            primaryRecipients.Add(bccList[0]);
                        }

                        if (sendEmail(primaryRecipients, bccList, emailSubject, emailBody, po.Attachments))
                        {
                            notifiedlist.Add(po.PurchaseOrderNumber);
                        }

                        /*
                                                log.Info("[MONITOR] email message: " + emailBody);
                                                if ((Mode == live) || (Mode == monitor))
                                                {
                                                    log.Info("[ProcessPOData] Mode: " + Mode.ToString() + ", buyer: " + po.Buyer + ", buyer email:" + buyerEmployee.EmailAddress);

                                                    NotifyEmployee(notifiedlist, po.PurchaseOrderNumber, buyerEmployee.EmailAddress, po.ReceivedBy, po.Bin, emailSubject, emailBody, po.Attachments);
                                                    if (projectManagerEmployee.EmailAddress.Length > 0 && buyerEmployee.EmailAddress != projectManagerEmployee.EmailAddress)
                                                    {
                                                        NotifyEmployee(notifiedlist, po.PurchaseOrderNumber, projectManagerEmployee.EmailAddress, po.ReceivedBy, po.Bin, emailSubject, emailBody, po.Attachments);
                                                    }
                                                }
                                                else
                                                {
                                                    log.Info("[ProcessPOData] Debug: Notification email would have been sent to buyer: " + po.Buyer + " and/or Project Manager: " + job.ProjectManagerName);
                                                }

                                                if (((Mode == monitor) || (Mode == debug)) &&
                                                    (MonitorEmailAddresses != null && MonitorEmailAddresses.Length > 0))
                                                {
                                                    foreach (var emailAddress in MonitorEmailAddresses)
                                                    {
                                                        if (sendEmail(emailAddress, emailSubject, emailBody, po.Attachments))
                                                        {
                                                            if (!notifiedlist.Contains(po.PurchaseOrderNumber))
                                                            {
                                                                notifiedlist.Add(po.PurchaseOrderNumber);
                                                            }
                                                        }
                                                    }
                                                }
                        */
                    }
                }
                catch (Exception ex)
                {
                    log.Info(ex.Message);
                }
/*                
                OdbcDataReader reader = cmd.ExecuteReader();
                try
                {
                    var EmployeeEmailAddresses = GetEmployees(cn);
                    var buyerColumn = reader.GetOrdinal("buyer");
                    var poNumColumn = reader.GetOrdinal("ponum");
                    var jobNumberColumn = reader.GetOrdinal("defaultjobnum");
                    var binColumn = reader.GetOrdinal("user_1");
                    var receivedByColumn = reader.GetOrdinal("user_2");
                    var receivedOnDateColumn = reader.GetOrdinal("user_3");
                    var vendorNameColumn = reader.GetOrdinal("vendorName");
                    var notesColumn = reader.GetOrdinal("user_6");
                    var workOrderNumberColumn = reader.GetOrdinal("defaultworkorder");
                    var attachmentIdColumn = reader.GetOrdinal("attachid");

                    while (reader.Read())
                    {
                        var purchaseOrderNumber = reader.GetString(poNumColumn);
                        var receivedBy = reader.GetString(receivedByColumn);
                        var bin = reader.GetString(binColumn);
                        var receivedOnDate = reader.GetDate(receivedOnDateColumn).ToShortDateString();
                        var buyer = reader.GetString(buyerColumn).ToLower();
                        var vendor = reader.GetString(vendorNameColumn);
                        var notes = reader.GetString(notesColumn);
                        var workOrderNumber = reader.GetString(workOrderNumberColumn);
                        var attachmentId = reader.GetString(attachmentIdColumn);
                        var job = queries.GetEmailBodyInformation(reader.GetString(jobNumberColumn), purchaseOrderNumber, workOrderNumber);
                        var buyerEmployee = GetEmployeeInformation(EmployeeEmailAddresses, buyer);
                        var projectManagerEmployee = job.ProjectManagerName.Length > 0 ? GetEmployeeInformation(EmployeeEmailAddresses, job.ProjectManagerName) : new Employee();
                        var attachments = queries.GetAttachmentsForPO(attachmentId);

                        log.Info("[ProcessPOData] ----------------- Found PO Number " + purchaseOrderNumber + " -------------------");

                        var emailSubject = String.Format(EmailSubject, purchaseOrderNumber, vendor);
                        var emailBody = FormatEmailBody(receivedOnDate, purchaseOrderNumber, receivedBy, bin, buyerEmployee.Name, vendor, job, notes);

                        log.Info("[MONITOR] email message: " + emailBody);
                        if ((Mode == live) || (Mode == monitor))
                        {
                            log.Info("[ProcessPOData] Mode: " + Mode.ToString() + ", buyer: " + buyer + ", buyer email:" + buyerEmployee.EmailAddress);

                            NotifyEmployee(notifiedlist, purchaseOrderNumber, buyerEmployee.EmailAddress, receivedBy, bin, emailSubject, emailBody, attachments);
                            if (projectManagerEmployee.EmailAddress.Length > 0 && buyerEmployee.EmailAddress != projectManagerEmployee.EmailAddress)
                            {
                                NotifyEmployee(notifiedlist, purchaseOrderNumber, projectManagerEmployee.EmailAddress, receivedBy, bin, emailSubject, emailBody, attachments);
                            }
                        }
                        else
                        {
                            log.Info("[ProcessPOData] Debug: Notification email would have been sent to buyer: " + buyer + " and/or Project Manager: " + job.ProjectManagerName);
                        }

                        if (((Mode == monitor) || (Mode == debug)) &&
                            (MonitorEmailAddresses != null && MonitorEmailAddresses.Length > 0))
                        {
                            foreach (var emailAddress in MonitorEmailAddresses)
                            {

                                if (sendEmail(emailAddress, emailSubject, emailBody, attachments))
                                {
                                    if (!notifiedlist.Contains(purchaseOrderNumber))
                                    {
                                        notifiedlist.Add(purchaseOrderNumber);
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception x)
                {
                    log.Error("[ProcessPOData] Reader Error: " + x.Message);
                }
                */
                foreach (string poNum in notifiedlist)
                {
                    try
                    {
                        dbRepository.MarkPOAsNotified(poNum);
                    }
                    catch (Exception x)
                    {
                        log.Error(String.Format("[ProcessPOData] Error updating PO {0} to be Notified: {1}", poNum, x.Message));
                    }
                }

                cn.Close();
            }
            catch (Exception x)
            {
                log.Error("[ProcessPOData] Exception: " + x.Message);
                return;
            }

            return;
        }

        public static Employee GetEmployeeInformation(List<Employee> EmployeeEmailAddresses, string employee)
        {
            try
            {
                var e = EmployeeEmailAddresses.FirstOrDefault(x => x.EmployeeId.ToLowerInvariant() == employee.ToLowerInvariant());

                if (e != null && e.EmailAddress.Length > 0)
                {
                    return e;
                }
            }
            catch (KeyNotFoundException)
            {
                log.Info("[GetEmployeeInformation] No Employee record found by employeeid for : " + employee);
            }
            catch (Exception x)
            {
                log.Error("[GetEmployeeInformation] by employeeid exception: " + x.Message);
            }

            try
            {
                var e = EmployeeEmailAddresses.FirstOrDefault(x => x.Name.ToLowerInvariant() == employee.ToLowerInvariant());

                if (e != null && e.EmailAddress.Length > 0)
                {
                    return e;
                }
            }
            catch (KeyNotFoundException)
            {
                log.Info("[GetEmployeeInformation] No Employee record found by name for : " + employee);
            }
            catch (Exception x)
            {
                log.Error("[GetEmployeeInformation] by name exception: " + x.Message);
            }

            return new Employee();
        }

        private static string FormatEmailBody(string receivedOnDate, string purchaseOrderNumber, string receivedBy, string bin, string buyerName, string vendor, JobInformation job, string notes)
        {
            var purchaseOrderItemTable = string.Empty;
            foreach (PurchaseOrderItem poItem in job.PurchaseOrderItems)
            {
                purchaseOrderItemTable += string.Format(messageBodyTableItem, poItem.ItemNumber, poItem.Description, poItem.Quantity);
            }

            var emailBody = String.Format(MessageBodyFormat, purchaseOrderNumber, receivedBy, receivedOnDate, bin, job.JobNumber, job.JobName, job.WorkOrderNumber, job.WorkOrderSite, job.CustomerName, vendor, buyerName, notes) + purchaseOrderItemTable + messageBodyTail;

            return emailBody;
        }

        private static bool sendEmail(ArrayList toEmailAddresses, ArrayList bccList, string subject, string emailBody, List<Attachment> poAttachments)
        {
            bool result = true;

            if (toEmailAddresses.Count == 0)
            {
                log.Error("  [sendEmail] No toEmailAddress to send message to");
                return false;
            }

            log.Info("  [sendEmail] Sending Email to: " + toEmailAddresses.Count);
            log.Info("  [sendEmail] subject: " + subject);
            log.Info("  [sendEmail] attachments: " + poAttachments.Count);

            try
            {
                using (MailMessage mail = new MailMessage())
                {
                    mail.From = new MailAddress(FromEmailAddress, "PO Notification");
                    foreach (string primaryEmailAddress in toEmailAddresses)
                    {
                        mail.To.Add(primaryEmailAddress);
                        log.Info("  [sendEmail] Sending email to primary address: " + primaryEmailAddress);
                    }

                    foreach (string bccEmailAddress in bccList)
                    {
                        mail.Bcc.Add(bccEmailAddress);
                        log.Info("  [sendEmail] Sending email to bcc address: " + bccEmailAddress);
                    }

                    mail.Subject = subject;
                    mail.Body = emailBody;
                    mail.IsBodyHtml = true;

                    foreach(Attachment poAttachment in poAttachments)
                    {
                        mail.Attachments.Add(poAttachment);
                    }

                    using (SmtpClient smtp = new SmtpClient(FromEmailSMTP, FromEmailPort.Value))
                    {
                        smtp.Credentials = new NetworkCredential(FromEmailAddress, FromEmailPassword);
                        smtp.EnableSsl = true;
                        smtp.Send(mail);
                        log.Info("  [sendEmail] Email Sent successfully");
                    }

                    mail.Attachments.Dispose();
                    mail.Dispose();
                }
            }
            catch (Exception x)
            {
                result = false;
                log.Error(String.Format("  [sendEmail] Error Sending email: {0}, message: {1}", x.InnerException, emailBody));
            }

            return result;
        }

    }
}
/*
log.Info("column names");
for (int col = 0; col < reader.FieldCount; col++)
{
    log.Info(reader.GetName(col));
}
*/
