using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Net;
using System.Net.Mail;
using JUST.Shared.DatabaseRepository;
using JUST.Shared.Classes;
using JUST.Shared.Utilities;

namespace JUST.PONotifier
{
    public class MainClass
    {
        /* version 1.03  */
        private const string debug = "debug";
        private const string live = "live";
        private const string monitor = "monitor";

        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private static ArrayList ValidModes = new ArrayList() { debug, live, monitor };
        private static Config config = new Config();

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

                config.getConfiguration(ValidModes);

                ProcessPOData();

                log.Info("[Main] Completion at " + DateTime.Now);
            }
            catch (Exception ex)
            {
                log.Error("[Main] Error: " + ex.Message);
            }
        }

        private static void ProcessPOData()
        {
            try
            {
                OdbcConnection cn;
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
                just.Add("Uid", config.Uid);
                just.Add("Pwd", config.Pwd);

                cn = new OdbcConnection(just.ConnectionString);
                cn.Open();
                log.Info("[ProcessPOData] Connection to database opened successfully");
                var dbRepository = new DatabaseRepository(cn, log, config.POAttachmentBasePath);
                
                List<PurchaseOrder> purchaseOrdersToNotify;
                try
                {
                    purchaseOrdersToNotify = dbRepository.GetPurchaseOrdersToNotify();
                    log.Info("purchaseOrdersToNotify found " + purchaseOrdersToNotify.Count.ToString() + " items.");

                    foreach (PurchaseOrder po in purchaseOrdersToNotify)
                    {
                        var job = dbRepository.GetEmailBodyInformation(po.JobNumber, po.PurchaseOrderNumber, po.WorkOrderNumber);
                        var buyerEmployee =  EmployeeLookup.FindEmployeeFromAllEmployees(dbRepository.GetEmployees(), po.Buyer);
                        var projectManagerEmployee = job.ProjectManagerName.Length > 0 ? EmployeeLookup.FindEmployeeFromAllEmployees(dbRepository.GetEmployees(), job.ProjectManagerName) : new Employee();

                        log.Info("[ProcessPOData] ----------------- Found PO Number " + po.PurchaseOrderNumber + " -------------------");

                        var emailSubject = String.Format(EmailSubject, po.PurchaseOrderNumber, po.Vendor);
                        var emailBody = FormatEmailBody(po.ReceivedOnDate, po.PurchaseOrderNumber, po.ReceivedBy, po.Bin, buyerEmployee.Name, po.Vendor, job, po.Notes);

                        ArrayList primaryRecipients = new ArrayList();
                        ArrayList bccList = new ArrayList();
                        if ((config.Mode == live) || (config.Mode == monitor))
                        {
                            primaryRecipients.Add(buyerEmployee.EmailAddress);
                            if (projectManagerEmployee.EmailAddress.Length > 0 && buyerEmployee.EmailAddress != projectManagerEmployee.EmailAddress)
                            {
                                primaryRecipients.Add(projectManagerEmployee.EmailAddress);
                            }
                        }

                        if (((config.Mode == monitor) || (config.Mode == debug)) &&
                            (config.MonitorEmailAddresses != null && config.MonitorEmailAddresses.Length > 0))
                        {
                            foreach(string monitorEmailAddress in config.MonitorEmailAddresses)
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
                    }
                }
                catch (Exception ex)
                {
                    log.Info(ex.Message);
                }

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
                    mail.From = new MailAddress(config.FromEmailAddress, "PO Notification");
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

                    using (SmtpClient smtp = new SmtpClient(config.FromEmailSMTP, config.FromEmailPort.Value))
                    {
                        smtp.Credentials = new NetworkCredential(config.FromEmailAddress, config.FromEmailPassword);
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
