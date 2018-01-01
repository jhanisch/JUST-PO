using System;
using System.Collections;
using System.Configuration;
//using System.Data;
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
        private static string DBConnectionString;
        private static string FromEmailAddress;
        private static string FromEmailPassword;
        private static string FromEmailSMTP;
        private static int FromEmailPort;
        private static string Mode;
        private static String MonitorEmailAddress;
        private static ArrayList ValidModes = new ArrayList() { debug, live, monitor };

        private static string MessageBodyFormat = "<h2>Purchase Order {0} Received</h2><br><h3>Purchase Order {0} has been received and has been placed in bin {1}</h3>";

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
            DBConnectionString = ConfigurationManager.ConnectionStrings["JUST"].ConnectionString;
            FromEmailAddress = ConfigurationManager.AppSettings["FromEmailAddress"];
            FromEmailPassword = ConfigurationManager.AppSettings["FromEmailPassword"];
            FromEmailSMTP = ConfigurationManager.AppSettings["FromEmailSMTP"];
            FromEmailPort = Convert.ToInt16(ConfigurationManager.AppSettings["FromEmailPort"]);

            Mode = ConfigurationManager.AppSettings["Mode"].ToLower();
            MonitorEmailAddress = ConfigurationManager.AppSettings["MonitorEmailAddress"];

            #region Validate Configuration Data
            var errorMessage = new StringBuilder();
            if (String.IsNullOrEmpty(FromEmailAddress))
            {
                errorMessage.Append("From Email Address is Required");
            }

            if (String.IsNullOrEmpty(FromEmailPassword))
            {
                errorMessage.Append("From Email Password is Required");
            }

            if (String.IsNullOrEmpty(FromEmailSMTP))
            {
                errorMessage.Append("From Email SMTP address is Required");
            }

            if (String.IsNullOrEmpty(FromEmailPassword))
            {
                errorMessage.Append("From Email Password is Required");
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
            string queryString =
                "SELECT po_num, descr FROM dbo.po where RCVD_DATE is not null and NOTIFIED is null;";
            ArrayList notified = new ArrayList();

            using (SqlConnection connection = new SqlConnection(
                DBConnectionString))
            {
                SqlCommand command = new SqlCommand(
                    queryString, connection);
                connection.Open();

                SqlDataReader reader = command.ExecuteReader();
                try
                {
                    while (reader.Read())
                    {
                        var subject = String.Format("PO {0} Received", reader[0]);
                        var message = String.Format(MessageBodyFormat, reader[0], reader[1]);
                        if ((Mode == live) || (Mode == monitor)) 
                        {
                            if (sendEmail("f1nut@aol.com", subject, message))
                            {
                                notified.Add(reader[0].ToString());
                            }
                        }

                        if (((Mode == monitor) || (Mode == debug)) && 
                            (!String.IsNullOrEmpty(MonitorEmailAddress))) 
                        {
                            if (sendEmail(MonitorEmailAddress, subject, message))
                            {
                                if (!notified.Contains(reader[0]))
                                {
                                    notified.Add(reader[0].ToString());
                                }
                            }
                        }

                    }
                }
                finally
                {
                    // Always call Close when done reading.
                    reader.Close();
                }

                foreach(var poNum in notified) {
                    try
                    {
                        SqlCommand update = new SqlCommand("update dbo.po SET notified='Y' WHERE po_num=@PoNum", connection);
                        update.Parameters.Add("@PoNum", SqlDbType.Int, 4).Value = poNum;
                        update.ExecuteNonQuery();
                    }
                    catch (Exception x)
                    {
                        log.Error(String.Format("Error updating PO as Notified: {0}", x.Message);
                    }
                }
            }
        }

        private static bool sendEmail(string toEmailAddress, string subject, string emailBody) {
            bool result = true;

            try
            {
                using (MailMessage mail = new MailMessage())
                {
                    mail.From = new MailAddress(FromEmailAddress);
                    mail.To.Add(toEmailAddress);
                    mail.Subject = subject;
                    mail.Body = emailBody;
                    mail.IsBodyHtml = true;

                    using (SmtpClient smtp = new SmtpClient(FromEmailSMTP, FromEmailPort))
                    {
                        smtp.Credentials = new NetworkCredential(FromEmailAddress, FromEmailPassword);
                        smtp.EnableSsl = true;
                        smtp.Send(mail);
                    }
                }
            }
            catch (Exception x) {
                result = false;
                log.Error(String.Format("Error Sending email to {0}, message: {1}", x.Message, emailBody));
            }

            return result;
        }
    }
}
