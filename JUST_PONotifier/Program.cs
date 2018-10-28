using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Odbc;
using System.Data.SqlClient;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using JUST.PONotifier.Classes;

namespace JUST.PONotifier
{
    class MainClass
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

        private static string EmailSubject = "Purchase Order {0} Received from {1}";
        private static string MessageBodyFormat = "<h2>Purchase Order {0} Received</h2><br><h3>Purchase Order {0} has been received by {1} and has been placed in bin {2}</h3>";
        private static string MessageBodyFormat2 = @"
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
        Vendor Name: {0}<br/>
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
                OdbcCommand cmd;
                string POQuery;
                var notifiedlist = new ArrayList();

                // user_1 = receiving rack location
                // user_2 = Receiver
                // user_3 = Received Date
                // user_4 = Bin Cleared Date
                // user_5 = Notified
                POQuery = "Select icpo.buyer, icpo.ponum, icpo.user_1, icpo.user_2, icpo.user_3, icpo.user_4, icpo.user_5, icpo.defaultjobnum, vendor.name as vendorName, icpo.user_6, icpo.defaultworkorder from icpo inner join vendor on vendor.vennum = icpo.vennum where icpo.user_3 is not null and icpo.user_5 = 0 order by icpo.ponum asc";

                OdbcConnectionStringBuilder just = new OdbcConnectionStringBuilder();
                just.Driver = "ComputerEase";
                just.Add("Dsn", "Company 0");
                just.Add("Uid", Uid);
                just.Add("Pwd", Pwd);

                cn = new OdbcConnection(just.ConnectionString);
                cmd = new OdbcCommand(POQuery, cn);
                cn.Open();
                log.Info("[ProcessPOData] Connection to database opened successfully");

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
                        var job = GetEmailBodyInformation(cn, reader.GetString(jobNumberColumn), purchaseOrderNumber, workOrderNumber);
                        var buyerEmployee = GetEmployeeInformation(EmployeeEmailAddresses, buyer);
                        var projectManagerEmployee = job.ProjectManagerName.Length > 0 ? GetEmployeeInformation(EmployeeEmailAddresses, job.ProjectManagerName) : new Employee();

                        log.Info("[ProcessPOData] ----------------- Found PO Number " + purchaseOrderNumber + " -------------------");

                        var emailSubject = String.Format(EmailSubject, purchaseOrderNumber, vendor);
                        var emailBody = FormatEmailBody(receivedOnDate, purchaseOrderNumber, receivedBy, bin, buyerEmployee.Name, vendor, job, notes);

                        log.Info("[MONITOR] email message: " + emailBody);
                        if ((Mode == live) || (Mode == monitor))
                        {
                            log.Info("[ProcessPOData] Mode: " + Mode.ToString() + ", buyer: " + buyer + ", buyer email:" + buyerEmployee.EmailAddress);

                            NotifyEmployee(notifiedlist, purchaseOrderNumber, buyerEmployee.EmailAddress, receivedBy, bin, emailSubject, emailBody);
                            if (projectManagerEmployee.EmailAddress.Length > 0 && buyerEmployee.EmailAddress != projectManagerEmployee.EmailAddress)
                            {
                                NotifyEmployee(notifiedlist, purchaseOrderNumber, projectManagerEmployee.EmailAddress, receivedBy, bin, emailSubject, emailBody);
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

                                if (sendEmail(emailAddress, emailSubject, emailBody))
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

                foreach (var poNum in notifiedlist)
                {
                    try
                    {
                        var updateCommand = string.Format("update icpo set \"user_5\" = 1 where icpo.ponum = '{0}'", poNum);
                        cmd = new OdbcCommand(updateCommand, cn);
                        cmd.ExecuteNonQuery();
                    }
                    catch (Exception x)
                    {
                        log.Error(String.Format("[ProcessPOData] Error updating PO {0} to be Notified: {1}", poNum, x.Message));
                    }
                }

                reader.Close();
                cn.Close();
            }
            catch (Exception x)
            {
                log.Error("[ProcessPOData] Exception: " + x.Message);
                return;
            }

            return;
        }

        private static Employee GetEmployeeInformation(List<Employee> EmployeeEmailAddresses, string employee)
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

            var emailBody = String.Format(MessageBodyFormat2, purchaseOrderNumber, receivedBy, receivedOnDate, bin, job.JobNumber, job.JobName, job.WorkOrderNumber, job.WorkOrderSite, job.CustomerName, vendor, buyerName, notes) + purchaseOrderItemTable + messageBodyTail;

            return emailBody;
        }

        private static void NotifyEmployee(ArrayList notifiedlist, string poNum, string employeeEmailAddress, string receivedBy, string bin, string emailSubject, string emailBody)
        {
            try
            {
                if (employeeEmailAddress.Length > 0)
                {
                    log.Info("  [NotifyEmployee]   sending email to: " + employeeEmailAddress);
                    if (sendEmail(employeeEmailAddress, emailSubject, emailBody))
                    {
                        notifiedlist.Add(poNum);
                    }
                }
                else
                {
                    log.Error("  [NotifyEmployee]  Purchase Order does not have an email address defined [" + emailSubject + "]");
                }
            }
            catch (Exception ex)
            {
                log.Info("  [NotifyEmployee] Error " + ex.Message);
            }
        }

        private static bool sendEmail(string toEmailAddress, string subject, string emailBody)
        {
            bool result = true;
            if (toEmailAddress.Length == 0)
            {
                log.Error("  [sendEmail] No toEmailAddress to send message to");
                return false;
            }

            log.Info("  [sendEmail] Sending Email to: " + toEmailAddress);

            try
            {
                using (MailMessage mail = new MailMessage())
                {
                    mail.From = new MailAddress(FromEmailAddress, "PO Notification");
                    mail.To.Add(toEmailAddress);
                    mail.Subject = subject;
                    mail.Body = emailBody;
                    mail.IsBodyHtml = true;

                    using (SmtpClient smtp = new SmtpClient(FromEmailSMTP, FromEmailPort.Value))
                    {
                        smtp.Credentials = new NetworkCredential(FromEmailAddress, FromEmailPassword);
                        smtp.EnableSsl = true;
                        smtp.Send(mail);
                        log.Info("  [sendEmail] Email Sent to " + toEmailAddress);
                    }
                }
            }
            catch (Exception x)
            {
                result = false;
                log.Error(String.Format("  [sendEmail] Error Sending email to {0}, message: {1}", x.Message, emailBody));
            }

            return result;
        }

        private static List<Employee> GetEmployees(OdbcConnection cn)
        {
            var employees = new List<Employee>();

            var buyerQuery = "Select user_1, user_2, name from premployee where user_1 is not null";
            var buyerCmd = new OdbcCommand(buyerQuery, cn);

            OdbcDataReader buyerReader = buyerCmd.ExecuteReader();

            while (buyerReader.Read())
            {
                var buyer = buyerReader.GetString(0);
                var email = buyerReader.GetString(1);
                var name = buyerReader.GetString(2);

                if (buyer.Trim().Length > 0)
                {
                    employees.Add(new Employee(buyer, name, email));
                }
            }

            buyerReader.Close();

            return employees;
        }

        private static JobInformation GetEmailBodyInformation(OdbcConnection cn, string jobNum, string purchaseOrderNumber, string workOrderNumber)
        {
            // from the jcjob table (Company 0)
            // user_1 = job description
            // user_2 = sales person
            // user_3 = designer
            // user_4 = Project Manager
            // user_5 = SM
            // user_6 = Fitter
            // user_7 = Plumber
            // user_8 = Tech 1
            // user_9 = Tech 2

            // add to email:
            //  user_1 
            var jobQuery = "Select jcjob.user_4, jcjob.user_1, customer.name, customer.cusNum from jcjob inner join customer on customer.cusnum = jcjob.cusnum where jobnum = '{0}'";
            var jobCmd = new OdbcCommand(string.Format(jobQuery, jobNum), cn);
            var result = new JobInformation();
            result.JobNumber = jobNum;

            OdbcDataReader jobReader = jobCmd.ExecuteReader();

            if (jobReader.Read())
            {
                result.ProjectManagerName = jobReader.GetString(0).ToLower();
                result.JobName = jobReader.GetString(1);
                result.CustomerName = jobReader.GetString(2);
                result.CustomerNumber = jobReader.GetString(3);
            }
            else
            {
                log.Info(string.Format("  [GetEmailBodyInformation] ERROR: Job Number (jobnum) {0} not found in jcjob", jobNum));
            }

            jobReader.Close();

            if (workOrderNumber.Trim().Length > 0)
            {
                var workOrderQuery = "Select dporder.workorder, dpsite.name from dporder inner join dpsite on dpsite.sitenum = dporder.sitenum where workorder = '{0}'";
                var workOrderCmd = new OdbcCommand(string.Format(workOrderQuery, workOrderNumber), cn);
                var workOrderResult = new JobInformation();

                OdbcDataReader workOrderReader = workOrderCmd.ExecuteReader();

                if (workOrderReader.Read())
                {
                    result.WorkOrderNumber = workOrderReader.GetString(0);
                    result.WorkOrderSite = workOrderReader.GetString(1);
                    log.Info(string.Format("  [GetEmailBodyInformation] INFO: Work Order Number {0} found, Site {1}", workOrderNumber, result.WorkOrderSite));
                }
                else
                {
                    log.Info(string.Format("  [GetEmailBodyInformation] INFO: Work Order Number (ponum) {0} not found in work order", workOrderNumber));
                }

                workOrderReader.Close();
            }

            var poIitemQuery = "Select icpoitem.itemnum, icpoitem.des, icpoitem.outstanding, icpoitem.unitprice from icpoitem where ponum = '{0}' order by icpoitem.itemnum asc";
            var poItemCmd = new OdbcCommand(string.Format(poIitemQuery, purchaseOrderNumber), cn);
            OdbcDataReader poItemReader = poItemCmd.ExecuteReader();

            while (poItemReader.Read())
            {
                result.PurchaseOrderItems.Add(new PurchaseOrderItem(poItemReader.GetString(0), poItemReader.GetString(1), poItemReader.GetString(2), poItemReader.GetString(3)));
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
