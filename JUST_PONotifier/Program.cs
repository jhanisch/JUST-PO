﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Odbc;
using System.Data.SqlClient;
using System.Net;
using System.Net.Mail;
using System.Text;

namespace JUST_PONotifier
{
    class MainClass
    {
        private const string debug = "debug";
        private const string live = "live";
        private const string monitor = "monitor";

        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private static string Uid;
        private static string Pwd;
        private static string DBConnectionString;
        private static string FromEmailAddress;
        private static string FromEmailPassword;
        private static string FromEmailSMTP;
        private static int? FromEmailPort;
        private static string Mode;
        private static String MonitorEmailAddress;
        private static ArrayList ValidModes = new ArrayList() { debug, live, monitor };

        private static string MessageBodyFormat = "<h2>Purchase Order {0} Received</h2><br><h3>Purchase Order {0} has been received by {1} and has been placed in bin {2}</h3>";

        static void Main(string[] args)
        {
            try {
                log.Info("Starting up at " + DateTime.Now);

                getConfiguration();

                ProcessPOData();

                log.Info("Completion at " + DateTime.Now);
            }
            catch (Exception ex)
            {
                log.Error("Error: " + ex.Message);
            }
        }

        private static void getConfiguration()
        {
            DBConnectionString = ConfigurationManager.ConnectionStrings["JUSTodbc"].ConnectionString;
            Uid = ConfigurationManager.AppSettings["Uid"];
            Pwd = ConfigurationManager.AppSettings["Pwd"];
            FromEmailAddress = ConfigurationManager.AppSettings["FromEmailAddress"];
            FromEmailPassword = ConfigurationManager.AppSettings["FromEmailPassword"];
            FromEmailSMTP = ConfigurationManager.AppSettings["FromEmailSMTP"];
            FromEmailPort = Convert.ToInt16(ConfigurationManager.AppSettings["FromEmailPort"]);

            Mode = ConfigurationManager.AppSettings["Mode"].ToLower();
            MonitorEmailAddress = ConfigurationManager.AppSettings["MonitorEmailAddress"];

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
                if (String.IsNullOrEmpty(MonitorEmailAddress))
                {
                    errorMessage.Append("Monitor Email Address is Required in monitor mode");
                }
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

                POQuery = "Select icpo.buyer, icpo.ponum, icpo.user_1, icpo.user_2, icpo.user_3, icpo.user_4, icpo.user_5, icpo.defaultjobnum from icpo where icpo.user_3 is not null and icpo.user_5 = 0 order by icpo.ponum asc";
                // "user = premployee"
                // user_1 = receiving rack location
                // user_2 = Receiver
                // user_3 = Received Date
                // user_4 = Bin Cleared Date
                // user_5 = Notified

                OdbcConnectionStringBuilder just = new OdbcConnectionStringBuilder();
                just.Driver = "ComputerEase";
                just.Add("Dsn", "Company 0");
                just.Add("Uid", Uid);
                just.Add("Pwd", Pwd);

                cn = new OdbcConnection(just.ConnectionString);
                cmd = new OdbcCommand(POQuery, cn);
                cn.Open();
                log.Info("connection opened successfully");

                OdbcDataReader reader = cmd.ExecuteReader();
                try
                {
                    var EmployeeEmailAddresses = GetEmployees(cn);
                    var buyerColumn = reader.GetOrdinal("buyer");
                    var poNumColumn = reader.GetOrdinal("ponum");
                    var jobNumberColumn = reader.GetOrdinal("defaultjobnum");
                    
                    while (reader.Read())
                    {
                        log.Info("Found PO Number " + reader.GetString(poNumColumn));

                        var buyer = reader.GetString(buyerColumn).ToLower();
                        var projectManager = GetPmForJob(cn, reader.GetString(jobNumberColumn)).ToLower();

                        if ((Mode == live) || (Mode == monitor))
                        {
                            NotifyEmployee(notifiedlist, EmployeeEmailAddresses, buyer, reader.GetString(poNumColumn), reader.GetString(3), reader.GetString(2));
                            NotifyEmployee(notifiedlist, EmployeeEmailAddresses, projectManager, reader.GetString(poNumColumn), reader.GetString(3), reader.GetString(2));

                        }
                        else
                        {
                            log.Info("Debug: Notification email would have been sent to buyer: " + buyer + " and/or Project Manager: " + projectManager);
                        }
/*
                        if (((Mode == monitor) || (Mode == debug)) &&
                            (!String.IsNullOrEmpty(MonitorEmailAddress)))
                        {
                            log.Info("sending email to monitor email address " + MonitorEmailAddress);
                            if (sendEmail(MonitorEmailAddress, subject, message))
                            {
                                if (!notifiedlist.Contains(reader.GetString(poNumColumn)))
                                {
                                    notifiedlist.Add(reader.GetString(poNumColumn));
                                }
                            }
                        }                       
*/
                    }
                }
                catch (Exception x)
                {
                    log.Error("Reader Error: " + x.Message);
                }

                foreach (var poNum in notifiedlist)
                {
                    try
                    {
                        log.Info("Updating po " + poNum);
/*                        var updateCommand = string.Format("update icpo set \"user_5\" = 1 where icpo.ponum = '{0}'", poNum);
                        log.Info("Updating po command: " + updateCommand.ToString());
                        cmd = new OdbcCommand(updateCommand, cn);
                        cmd.ExecuteNonQuery();
*/
                    }
                    catch (Exception x)
                    {
                        log.Error(String.Format("Error updating PO {0} to be Notified: {1}", poNum, x.Message));
                    }
                }

                cn.Close();
                log.Info("connection closed successfully");
            }
            catch (Exception x)
            {
                log.Error("Exception: " + x.Message);
                return;
            }

            return;
        }

        private static void NotifyEmployee(ArrayList notifiedlist, Dictionary<string, string> EmployeeEmailAddresses, string employee, string poNum, string receivedBy, string bin)
        {
            var subject = String.Format("Purchase Order {0} Received", poNum);
            var message = String.Format(MessageBodyFormat, poNum, receivedBy, bin);

            try
            {
                var emailAddress = EmployeeEmailAddresses[employee];

                log.Info("sending email to: " + emailAddress);

                if (emailAddress.Length > 0)
                {
                    if (sendEmail(emailAddress, subject, message))
                    {
                        notifiedlist.Add(poNum);
                    }
                }
                else
                {
                    log.Error("Purchase Order " + poNum + " does not have an email address defined");
                }
            }
            catch (Exception)
            {
                log.Info("No Email address found for " + employee);
            }
        }

        private static bool sendEmail(string toEmailAddress, string subject, string emailBody) {
            bool result = true;
            log.Info("Sending Email to: " + toEmailAddress);
            log.Info("  Message: " + emailBody);

            return true;

            try
            {
                using (MailMessage mail = new MailMessage())
                {
                    log.Info("  Using 1");
                    mail.From = new MailAddress(FromEmailAddress, "PO Notification");
                    mail.To.Add(toEmailAddress);
                    mail.Subject = subject;
                    mail.Body = emailBody;
                    mail.IsBodyHtml = true;

                    using (SmtpClient smtp = new SmtpClient(FromEmailSMTP, FromEmailPort.Value))
                    {
                        log.Info("  Using 2");
                        smtp.Credentials = new NetworkCredential(FromEmailAddress, FromEmailPassword);
                        smtp.EnableSsl = true;
                        smtp.Send(mail);
                        log.Info("  Email Sent");
                    }
                }
            }
            catch (Exception x) {
                result = false;
                log.Error(String.Format("Error Sending email to {0}, message: {1}", x.Message, emailBody));
            }

            return result;
        }

        private static Dictionary<string, string> GetEmployees(OdbcConnection cn)
        {
            var buyerQuery = "Select user_1, user_2 from premployee where user_1 is not null";
            var buyers = new Dictionary<string, string>();
            var buyerCmd = new OdbcCommand(buyerQuery, cn);

            OdbcDataReader buyerReader = buyerCmd.ExecuteReader();

            while (buyerReader.Read())
            {
                var buyer = buyerReader.GetString(0);
                var email = buyerReader.GetString(1);
                if (buyer.Trim().Length > 0)
                {
                    buyers.Add(buyer.ToLower(), email);
                }
            }

            return buyers;
        }

        private static string GetPmForJob(OdbcConnection cn, string jobNum)
        {
            // from the jcjob table (Company 0)
            // user_1 = job description
            // user_2 = sales person
            // user_3 = designer
            // user_4 = Projet Manager
            // user_5 = SM
            // user_6 = Fitter
            // user_7 = Plumber
            // user_8 = Tech 1
            // user_9 = Tech 2

            var jobQuery = "Select user_4 from jcjob where jobnum = '{0}'";
            var jobCmd = new OdbcCommand(string.Format(jobQuery, jobNum), cn);
            string tech = string.Empty;

            OdbcDataReader jobReader = jobCmd.ExecuteReader();

            while (jobReader.Read())
            {
                tech = jobReader.GetString(0);
            }

            return tech;
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
