using JUST.Shared.Classes;
using JUST.Shared.DatabaseRepository;
using JUST.Shared.Utilities;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data.Odbc;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Just.AgedReceivables
{
    public class NotificationType
    {
        public string PreFix { get; set; }
        public ArrayList NotificationList { get; set; }

        public NotificationType()
        {
            PreFix = String.Empty;
            NotificationList = new ArrayList();
        }

        public NotificationType(string preFix, ArrayList notificationList)
        {
            PreFix = preFix;
            NotificationList = new ArrayList();
            if (notificationList.Count > 0)
            {
                NotificationList.AddRange(notificationList);
            }
        }

    }

    public class AgedReceivablesConfig : Config
    {
        public long ThresholdDays;
        public ArrayList Notifications;

        public AgedReceivablesConfig(bool modeRequired) : base(modeRequired)
        {
            ThresholdDays = 60;
            Notifications = new ArrayList();
        }

        public void LoadConfiguration(ArrayList validModes, bool modeRequired)
        {
            var days = ConfigurationManager.AppSettings["ThresholdDays"];
            if (days != null && days.Length > 0)
            {
                try
                {
                    ThresholdDays = Convert.ToInt32(days);
                }
                catch
                {
                    ThresholdDays = 60;
                }
            }

            base.getConfiguration(validModes, modeRequired);

            if (HVACEmailAddresses != null && HVACEmailAddresses.Count > 0)
            {
                Notifications.Add(new NotificationType("H", HVACEmailAddresses));
            }

            if (PlumbingEmailAddresses != null && PlumbingEmailAddresses.Count > 0)
            {
                Notifications.Add(new NotificationType("P", PlumbingEmailAddresses));
            }

            if (MonitorEmailAddresses != null && MonitorEmailAddresses.Count > 0)
            {
                Notifications.Add(new NotificationType(String.Empty, MonitorEmailAddresses));
            }

        }
    }

    public class AgedReceivables
    {
        private const string debug = "debug";
        private const string live = "live";
        private const string monitor = "monitor";

        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private static ArrayList ValidModes = new ArrayList() { debug, live, monitor };
        private static AgedReceivablesConfig config = new AgedReceivablesConfig(false);
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
              <th colspan=""12"" style=""background-color: cyan; padding: 15px; text-align: left; font-size:24px""><strong>Aged Receivables - {0}</strong><span style=""font-size: 14px""><br/>> {1} days old</span><th>
           </tr>
           ";

        private static string EmailBodyHeaderLine = @"
           <tr style=""opacity: 0.75"">
              <td></td>
              <td colspan=""2"" style=""text-align: right; padding: 5px 10px 5px 5px"">Invoice #</td>
              <td colspan=""2"" style=""text-align: right; padding-left: 10px "">Job/WO #</td>
              <td colspan=""2"" style=""text-align: right; padding: 5px 10px 5px 5px"">Invoice Date</td>
              <td colspan=""2"" style=""text-align: right; padding: 5px 10px 5px 5px"">Due Date</td>
              <td colspan=""2"" style=""text-align: right; padding: 5px 10px 5px 5px"">Amount Due</td>
           </tr>";

        private static string EmailBodyTableHr = @"
           <tr>
              <td colspan=""12""><hr></td>
           </tr>";

        private static string EmailBodyTableCustomerLine = @"
           <tr>
              <td></td>
              <td colspan=""11"" style=""text-align: left; font-size: 20px""><strong>{0} {1}</strong></td>
           </tr>";

        private static string EmailBodyDetailLine = @"
           <tr>
              <td></td>
              <td colspan=""2"" style=""text-align: right; padding: 5px 10px 5px 5px"">{0}</td>
              <td colspan=""2"" style=""text-align: right; padding-left: 10px "">{1}</td>
              <td colspan=""2"" style=""text-align: right; padding: 5px 10px 5px 5px"">{2}</td>
              <td colspan=""2"" style=""text-align: right; padding: 5px 10px 5px 5px"">{3}</td>
              <td colspan=""2"" style=""text-align: right; padding: 5px 10px 5px 5px"">{4}</td>
           </tr>";

        private static string EmailBodyTotalLine = @"
           <tr>
              <td></td>
              <td></td>
              <td></td>
              <td></td>
              <td></td>
              <td></td>
              <td></td>
              <td></td>
              <td></td>
              <td colspan=""2"" style=""text-align: right; padding: 5px 10px 5px 5px""><hr></td>
           </tr>
           <tr>
              <td></td>
              <td></td>
              <td></td>
              <td></td>
              <td></td>
              <td></td>
              <td></td>
              <td></td>
              <td></td>
              <td colspan=""2"" style=""text-align: right; padding: 5px 10px 5px 5px"">{0}</td>
           </tr>";


        private static string EmailBodyMemoLine = @"
           <tr>
              <td></td>
              <td></td>
              <td></td>
              <td colspan=""9"" style=""color:grey; padding: 0px 0px 10px 5px; opacity: 0.75"">{0}</td>
           </tr>";


        private static string EmailBodyFooter = @"</table></body>";


        static void Main(string[] args)
        {
            try
            {
                log.Info("[AgedReceivables] Starting up at " + DateTime.Now);

                //                config.getConfiguration(ValidModes, false);
                config.LoadConfiguration(ValidModes, false);

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
            var lastCustomerNumber = String.Empty;
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
            var dbRepository = new DatabaseRepository(cn, log, null);

            var allOpenInvoices = dbRepository.GetAgedReceivables();
            var emailSubject = string.Format(EmailSubject, DateTime.Now.ToShortDateString());
            log.Info("[ProcessAgedReceivables] Found " + allOpenInvoices.Count() + " open invoices to process.");

            foreach (NotificationType n in config.Notifications)
            {
                var emailBody = string.Format(EmailBodyHeader, DateTime.Now.ToShortDateString(), config.ThresholdDays);
                log.Info("[ProcessAgedReceivables] " + n.PreFix);

                var agedReceivables = from invoice in allOpenInvoices
                                      where invoice.DueDate < DateTime.Now.AddDays(-1 * Math.Abs(config.ThresholdDays)) && (n.PreFix.Length > 0 ? invoice.WorkOrderNumber.StartsWith(n.PreFix) : !invoice.Notified)
                                      orderby invoice.CustomerName ascending, invoice.CustomerNumber ascending, invoice.InvoiceDate ascending
                                      select invoice;

                var grandTotal = 0.00M;
                var subTotal = 0.00M;

                if (agedReceivables.Count() > 0)
                {
                    log.Info("[ProcessAgedReceivables] Found " + agedReceivables.Count() + " aged invoices >= " + config.ThresholdDays.ToString() + " days old to notify.");
                    emailBody += EmailBodyTableHr + EmailBodyHeaderLine;

                    foreach (AgedReceivable ar in agedReceivables)
                    {
                        if (ar.CustomerNumber != lastCustomerNumber)
                        {
                            if (subTotal > 0)
                            {
                                emailBody += String.Format(EmailBodyTotalLine, subTotal.ToString("C"));
                            }

                            lastCustomerNumber = ar.CustomerNumber;
                            emailBody += EmailBodyTableHr;
                            emailBody += string.Format(EmailBodyTableCustomerLine, ar.CustomerName, "(" + ar.CustomerNumber + ")");

                            subTotal = 0.00M;
                        }

                        emailBody += string.Format(EmailBodyDetailLine, ((ar.DaysOverdue <= (Math.Abs(config.ThresholdDays) + 7)) ? "*" : String.Empty) + ar.InvoiceNumber, (ar.JobNumber.Length > 0 ? ar.JobNumber : ar.WorkOrderNumber), ar.InvoiceDate.ToShortDateString(), ar.DueDate.ToShortDateString(), ar.AmountDue.ToString("C"));
                        if (ar.Memo.Trim().Length > 0)
                        {
                            emailBody += string.Format(EmailBodyMemoLine, ar.Memo);
                        }

                        subTotal += ar.AmountDue;
                        grandTotal += ar.AmountDue;
                        ar.Notified = true;
                    }

                    // prints the subtotal for the last customer
                    if (subTotal > 0)
                    {
                        emailBody += String.Format(EmailBodyTotalLine, subTotal.ToString("C"));
                    }

                    emailBody += EmailBodyTableHr;
                }
                else
                {
                    emailBody += string.Format(EmailBodyTableCustomerLine, NoAgedReceivablesFound, string.Empty);
                }

                emailBody += string.Format(EmailBodyTotalLine, grandTotal.ToString("C"));
                emailBody += EmailBodyFooter;
                Utils.sendEmail(config, n.NotificationList, emailSubject, emailBody);
            }

            cn.Close();

            return;
        }

    }
}
