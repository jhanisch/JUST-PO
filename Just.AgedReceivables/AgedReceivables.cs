using JUST.Shared.Classes;
using JUST.Shared.DatabaseRepository;
using JUST.Shared.Utilities;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Just.AgedReceivables
{
    public class AgedReceivables
    {
        private const string debug = "debug";
        private const string live = "live";
        private const string monitor = "monitor";

        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private static ArrayList ValidModes = new ArrayList() { debug, live, monitor };
        private static Config config = new Config(false);
        private static string RunDate = DateTime.Now.ToShortDateString();
        private static string NoAgedReceivablesFound = "No Aged Receivables Found\n";

        private static string EmailSubject = @"Aged Receivables - {0}";
        private static string EmailBodyHeader = @"<body style = ""margin-left: 20px; margin-right:20px"" >
        <table style = ""width:100%"">
           <tr>
              <th style=""width:8.33%""></th>
              <th style=""width:8.33%""></th>
              <th style=""width:8.33%""></th>
              <th style=""width:8.33%""></th>
              <th style=""width:8.33%""></th>
              <th style=""width:8.33%""></th>
              <th style=""width:8.33%""></th>
              <th style=""width:8.33%""></th>
              <th style=""width:8.33%""></th>
              <th style=""width:8.33%""></th>
              <th style=""width:8.33%""></th>
              <th style=""width:8.33%""></th>
           </tr>
           <tr>
              <th colspan=""12"" style=""background-color: cyan; padding: 15px; text-align: left; font-size:24px""><strong>Aged Receivables - {0}</strong><th>
           </tr>
           ";

        private static string EmailBodyTableHr = @"
           <tr>
              <td colspan=""12""><hr></td>
           </tr>";

        private static string EmailBodyTableCustomerLine = @"
           <tr>
              <td></td>
              <td colspan=""11"" style=""text-align: left; font-size: 20px""><strong>{0}</strong></td>
           </tr>";

        private static string EmailBodyDetailLine = @"
           <tr>
              <td></td>
              <td colspan=""2"" style=""text-align: right; padding: 5px 10px 5px 5px"">{0}</td>
              <td colspan=""3"" style=""text-align: left; padding-left: 10px "">{1}</td>
              <td colspan=""2"" style=""text-align: right; padding: 5px 10px 5px 5px"">{2}</td>
              <td colspan=""3"" style=""text-align: left; padding-left: 10px"">{3}</td>
           </tr>";

        private static string EmailBodyFooter = @"</table></body>";


        static void Main(string[] args)
        {
            try
            {
                log.Info("[AgedReceivables] Starting up at " + DateTime.Now);

                config.getConfiguration(ValidModes, false);

                ProcessAgedReceivables();

                log.Info("[AgedReceivables] Completion at " + DateTime.Now);
            }
            catch (Exception ex)
            {
                log.Error("[AgedReceivables] Error: " + ex.Message);
            }
        }

        private static void ProcessAgedReceivables()
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
            log.Info("[ProcessAgedReceivables] Connection to database opened successfully");
            var dbRepository = new DatabaseRepository(cn, log, null);

            var agedReceivables = dbRepository.GetAgedReceivables();
            var emailSubject = string.Format(EmailSubject, DateTime.Now.ToShortDateString());
            var emailBody = string.Format(EmailBodyHeader, DateTime.Now.ToShortDateString());
            log.Info("[ProcessAgedReceivables] Found " + agedReceivables.Count() + " aged receivables to notify.");

            if (agedReceivables.Count() > 0)
            {
                emailBody += EmailBodyTableHr;

                foreach (AgedReceivable ar in agedReceivables)
                {
                    emailBody += string.Format(EmailBodyTableCustomerLine, ar.CustomerName);
                    emailBody += string.Format(EmailBodyDetailLine, "Invoice #:", ar.InvoiceNumber, ar.InvoiceDate, ar.AgedAmount.ToString("C"));
                    emailBody += EmailBodyTableHr;
                }
            }
            else
            {
                emailBody += string.Format(EmailBodyTableCustomerLine, NoAgedReceivablesFound);
            }

            emailBody += EmailBodyFooter;
            Utils.sendEmail(config, config.MonitorEmailAddresses, emailSubject, emailBody);
            cn.Close();

            return;
        }

    }
}
