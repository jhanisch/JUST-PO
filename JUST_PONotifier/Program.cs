using System;
using System.Configuration;
using System.Data.SqlClient;
using System.Net;
using System.Net.Mail;

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

        static void Main(string[] args)
        {
            try {
                log.Info("starting up at " + DateTime.Now);

                getDBParameters();

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
                "SELECT * FROM dbo.icPo where received date is not null and sent is null;";
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
                        var message = String.Format("Purchase Order {0} has been received and has been placed in bin {1}", reader[1], reader[2]);
                        sendEmail(reader[3].ToString(), message);

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
            var smtp = new SmtpClient(FromEmailSMPT);
            smtp.EnableSsl = true;
            smtp.Port = FromEmailPort;
            smtp.Credentials = new NetworkCredential(FromEmailAddress, FromEmailPassword);
            smtp.Send(FromEmailAddress, toEmailAddress, "PO Received", message);            
        }
    }
}
