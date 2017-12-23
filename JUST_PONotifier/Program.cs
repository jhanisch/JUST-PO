using log4net;
using System;
using System.Configuration;

namespace JUST_PONotifier
{
    class MainClass
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        private static string DBUserName;
        private static string DBPassword;
        private static string DBDsn;

        static void Main(string[] args)
        {
            log.Info("starting up at " + DateTime.Now);

            getDBParameters();


            log.Info("Completion at " + DateTime.Now);
        }

        private static void getDBParameters() {
            try {
                DBUserName = ConfigurationSettings.AppSettings["DBUserName"];
                DBPassword = ConfigurationSettings.AppSettings["DBPassword"];
                DBDsn = ConfigurationSettings.AppSettings["DBDsn"];

                Console.WriteLine("DB: " + DBUserName);
            }
            catch (Exception)
            {
                log.Error("Error reading app.config settings.  DBUserName: " + DBUserName + ", DBPassword: " + DBPassword + ", DBDsn: " + DBDsn);
            }
        }
    }
}
