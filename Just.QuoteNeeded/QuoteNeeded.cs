using JUST.Shared.Classes;
using JUST.Shared.DatabaseRepository;
using JUST.Shared.Utilities;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace Just.QuoteNeeded
{
    public class QuoteNeeded
    {
        private const string debug = "debug";
        private const string live = "live";
        private const string monitor = "monitor";

        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private static ArrayList ValidModes = new ArrayList() { debug, live, monitor };
        private static Config config = new Config(false);
        private static string RunDate = DateTime.Now.ToShortDateString();
        private static string DefaultHVACMessage = "No new HVAC Work Tickets requiring a quote were found.";
        private static string DefaultPlumbingMessage = "No new Plumbing Work Tickets requiring a quote were found.";

        private static string EmailSubject = @"Quotes Needed - {0}";
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
              <th colspan=""12"" style=""background-color: cyan; padding: 15px; text-align: left; font-size:24px""><strong>Quotes Needed - {0}</strong><th>
           </tr>
           ";

        private static string EmailBodyTableWorkOrderLine = @"
           <tr>
              <td></td>
              <td colspan=""11"" style=""text-align: left; font-size: 20px""><strong>{0} {1}</strong></td>
           </tr>";

        private static string EmailBodyTableDoubleLineItem = @"
           <tr>
              <td></td>
              <td colspan=""2"" style=""text-align: right; padding: 5px 10px 5px 5px""><strong>{0}:</strong></td>
              <td colspan=""3"" style=""text-align: left; padding-left: 10px "">{1}</td>
              <td colspan=""2"" style=""text-align: right; padding: 5px 10px 5px 5px""><strong>{2}:</strong></td>
              <td colspan=""3"" style=""text-align: left; padding-left: 10px"">{3}</td>
           </tr>";

        private static string EmailBodyTableSingleLineItem = @"
           <tr>
              <td></td>
              <td colspan=""2"" style=""text-align: right; padding: 5px 10px 5px 5px""><strong>{0}:</strong></td>
              <td colspan=""9"" style=""text-align: left; padding-left: 10px "">{1}</td>
           </tr>";

        private static string EmailBodyTableHr = @"
           <tr>
              <td colspan=""12""><hr></td>
           </tr>";


        private static string EmailBodyFooter = @"</table></body>";

        static void Main(string[] args)
        {
            try
            {
                log.Info("[QuoteNeeded] Starting up at " + DateTime.Now);

                config.getConfiguration(ValidModes, false);

                ProcessQuotesNeeded();

                log.Info("[QuoteNeeded] Completion at " + DateTime.Now);
            }
            catch (Exception ex)
            {
                log.Error("[QuoteNeeded] Error: " + ex.Message);
            }
        }

        private static void ProcessQuotesNeeded()
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
            log.Info("[ProcessQuotesNeeded] Connection to database opened successfully");
            var dbRepository = new DatabaseRepository(cn, log, null);


            var x = dbRepository.GetQuotesNeeded();
            log.Info("[ProcessQuotesNeeded] Found " + x.Count.ToString() + " quotes to notify");
            var hvacQuotes = x.FindAll(q => q.WorkOrder.StartsWith("H"));
            var plumbingQuotes = x.FindAll(q => !q.WorkOrder.StartsWith("H"));

            log.Info("[ProcessQuotesNeeded] Found " + hvacQuotes.Count.ToString() + " HVAC quotes to notify");
            log.Info("[ProcessQuotesNeeded] Found " + plumbingQuotes.Count.ToString() + " Plumbing quotes to notify");

            var emailSubject = string.Format(EmailSubject, RunDate);
            if (Utils.sendEmail(config, config.PlumbingEmailAddresses, emailSubject, FormatEmailBody(plumbingQuotes, dbRepository, DefaultPlumbingMessage)))
            {
                foreach(Quote quote in hvacQuotes)
                {
                    dbRepository.MarkQuoteAsNotified(quote);
                }
            }

            if (Utils.sendEmail(config, config.HVACEmailAddresses, emailSubject, FormatEmailBody(hvacQuotes, dbRepository, DefaultHVACMessage)))
            {
                foreach(Quote quote in plumbingQuotes)
                {
                    dbRepository.MarkQuoteAsNotified(quote);
                }
            }
        }

        private static string FormatEmailBody(List<Quote> quotes, DatabaseRepository dbRepository, string defaultMessage)
        {
            string date = DateTime.Now.ToShortDateString();

            var emailBody = string.Format(EmailBodyHeader, RunDate);

            if (quotes.Count == 0)
            {
                emailBody += string.Format(EmailBodyTableWorkOrderLine, defaultMessage, String.Empty);
                emailBody += EmailBodyFooter;
                return emailBody;
            }

            long row = 1;
            emailBody += EmailBodyTableHr;

            foreach (Quote quote in quotes)
            {
                emailBody += string.Format(EmailBodyTableWorkOrderLine, "Work Order", quote.WorkOrder);
                //, quote.WorkTicket, quote.CustomerName, quote.SiteName, quote.DescriptionOfWork, quote.TicketNote);
//                log.Info("[FormatEmailBody 3] " + emailBody);

                emailBody += string.Format(EmailBodyTableDoubleLineItem, "Work Ticket", quote.WorkTicket, "Site Name", quote.SiteName);
                emailBody += string.Format(EmailBodyTableDoubleLineItem, "Manaufactuer/Model", quote.Manufacturer + " / " + quote.Model, "Serial Number", quote.SerialNumber);
                emailBody += (quote.ServiceTech.Length > 0) ? string.Format(EmailBodyTableSingleLineItem, "Service Person", EmployeeLookup.FindEmployeeFromAllEmployees(dbRepository.GetEmployees(), quote.ServiceTech).Name) : string.Empty;
                emailBody += (quote.DescriptionOfWork.Length > 0) ? string.Format(EmailBodyTableSingleLineItem, "Original Description", quote.DescriptionOfWork) : string.Empty;            
                emailBody += (quote.TicketNote.Length > 0) ? string.Format(EmailBodyTableSingleLineItem, "Ticket Note", quote.TicketNote) : string.Empty;

                emailBody += EmailBodyTableHr;
                row++;
            }

            emailBody += EmailBodyFooter;
//            log.Info(emailBody);

            return emailBody;

        }
        /*
        private static bool sendEmail(string toEmailAddress, string subject, string emailBody)
        {

            log.Info(emailBody);

            bool result = true;
            if (toEmailAddress.Length == 0)
            {
                log.Error("  [sendEmail] No toEmailAddress to send message to");
                return false;
            }

            log.Info("  [sendEmail] Sending Email to: " + toEmailAddress);
            log.Info("  [sendEmail] EmailMessage: " + emailBody);

            try
            {
                using (MailMessage mail = new MailMessage())
                {
                    mail.From = new MailAddress(config.FromEmailAddress, "New Job Notification");
                    mail.To.Add(toEmailAddress);
                    mail.Subject = subject;
                    mail.Body = emailBody;
                    mail.IsBodyHtml = true;

                    using (SmtpClient smtp = new SmtpClient(config.FromEmailSMTP, config.FromEmailPort.Value))
                    {
                        smtp.Credentials = new NetworkCredential(config.FromEmailAddress, config.FromEmailPassword);
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
        */
    }
}
