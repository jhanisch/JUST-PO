using System;
using System.Configuration;
using System.Data.SqlClient;
using System.Net;
using System.Net.Mail;
using System.Security.Cryptography;

namespace JUST_PONotifier
{
    class MainClass
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private static string DBConnectionString;
        private static string FromEmailAddress;
        private static string FromEmailPassword;
        private static string FromEmailSMPT;
        private static int FromEmailPort;

        private static string MessageBodyFormat = "<h1>PO {0} Received</h1><br><h2>Purchase Order {0} has been received and has been placed in bin {1}</h2>";

        static void Main(string[] args)
        {
            try {
                log.Info("starting up at " + DateTime.Now);

                getDBParameters();

                ProcessPOData();

                log.Info("Completion at " + DateTime.Now);
            }
            catch (Exception ex)
            {
                log.Error("Error: " + ex.Message);
            }
        }

        private static void getDBParameters() {
            DBConnectionString = ConfigurationManager.ConnectionStrings["JUST"].ConnectionString;
            Console.WriteLine("DB: " + DBConnectionString);

            FromEmailAddress = ConfigurationManager.AppSettings["FromEmailAddress"];
            FromEmailPassword = ConfigurationManager.AppSettings["FromEmailPassword"];
            FromEmailSMPT = ConfigurationManager.AppSettings["FromEmailSMPT"];
            FromEmailPort = Convert.ToInt16(ConfigurationManager.AppSettings["FromEmailPort"]);
        }

        private static void ProcessPOData()
        {
            string queryString =
                "SELECT po_num, descr FROM dbo.po where RCVD_DATE is not null and NOTIFIED is null;";
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
                        var message = String.Format(MessageBodyFormat, reader[0], reader[1]);
                        sendEmail("f1nut@aol.com", message);

                        //Mark PO Record as notified
                    }
                }
                finally
                {
                    // Always call Close when done reading.
                    reader.Close();
                }
            }
        }

        private static void sendEmail(string toEmailAddress, string message) {
            try
            {
                using (MailMessage mail = new MailMessage())
                {
                    mail.From = new MailAddress(FromEmailAddress);
                    mail.To.Add(toEmailAddress);
                    mail.Subject = "PO Received";
                    mail.Body = message;
                    mail.IsBodyHtml = true;

                    using (SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587))
                    {
                        smtp.Credentials = new NetworkCredential(FromEmailAddress, FromEmailPassword);
                        smtp.EnableSsl = true;
                        smtp.Send(mail);
                    }
                }
            }
            catch (Exception x) {
                log.Error(x.Message);
            }
        }
    }
}
