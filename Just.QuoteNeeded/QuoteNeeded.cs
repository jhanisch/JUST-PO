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
        private static Config config = new Config();

        private static string EmailSubject = "Quotes Needed";
        private static string EmailBodyHeader = @"<body style = ""margin-left: 20px; margin-right:20px"" >
        <hr/>
        <h2>Quotes Needed</h2>
        <hr/>

        <table style = ""width:100%; text-align: left"" border=""1"" cellpadding=""10"" cellspacing=""0"">
          <thead>
            <tr style = ""background-color: cyan"" >
                <th>Work Order</th>
                <th>Work Ticket</th>
                <th>Customer Name</th>
                <th>Site</th>
                <th>Original Call Description</th>
                <th>Notes</th>
            </tr>
           </thead>
           <tbody>";
        private static string EmailBodyTableLineItem = @"<tr>
                <td>{0}</td>
                <td>{1}</td>
                <td>{2}</td>
                <td>{3}</td>
                <td>{4}</td>
                <td>{5}</td>
            </tr>";
        private static string EmailBodyFooter = @"</tbody></table></body>";

        private static string EmailBodyHeader2 = @"
<html lang=""en"">
<head>
<meta charset = ""utf-8"" >
<title>Quotes Needed</title>

<style>
* {
  box-sizing: border-box;
}

.row::after {
  content: """";
  clear: both;
  display: table;
}

[class*=""col-""] {
  float: left;
  padding: 15px;
}

.col-1 {width: 8.33%; padding: 5px}
.col-2 {width: 16.66%; padding: 5px}
.col-3 {width: 25%; padding: 5px}
.col-4 {width: 33.33%; padding: 5px}
.col-5 {width: 41.66%; padding: 5px}
.col-6 {width: 50%; padding: 5px}
.col-7 {width: 58.33%; padding: 5px}
.col-8 {width: 66.66%; padding: 5px}
.col-9 {width: 75%; padding: 5px}
.col-10 {width: 83.33%; padding: 5px}
.col-11 {width: 91.66%; padding: 5px}
.col-12 {width: 100%; padding: 5px}

.label { text-align: right }

html {
  font-family: ""Lucida Sans"", sans-serif;
}

.header {
  background-color: cyan;
  color: black;
  padding: 15px;
  border: 1px solid #dddddd
}

</style>
</head>

<body>

<div style=""background-color: cyan; color: black; padding: 15px; border: 1px solid #dddddd"">
<h2>Quotes Needed {0}</h2>
</div>

<!-- https://www.w3schools.com/css/css_rwd_grid.asp -->
<!-- https://www.w3schools.com/css/tryit.asp?filename=tryresponsive_styles -->";

        private static string EmailBodyTableItemBegin = @"
<div style=""{0}"">
	<div class=""row"">
	  <div style=""width: 100%; padding: 5px"">Work Order # {1}</div>
	</div>";

        private static string EmailBodyTableInnerItem2 = @"	<div class=""row"">
	  <div style=""width: 8.33%; padding: 5px""></div>
	  <div style=""width: 25%; padding: 5px"">{0}</div>
	  <div style=""width: 66.66%; padding: 5px"">{1}</div>
    </div>";

        private static string EmailBodyTablItemEnd = @"</div>";


        private static string EmailBodyFooter2 = @"</body></html>";

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

            var emailBody = FormatEmailBody(plumbingQuotes);
            log.Info(emailBody);

            Utils.sendEmail(config, "notifications@justserviceinc.com", EmailSubject, FormatEmailBody(plumbingQuotes));
            Utils.sendEmail(config, "notifications@justserviceinc.com", EmailSubject, FormatEmailBody(hvacQuotes));
        }

        private static string FormatEmailBody(List<Quote> quotes)
        {
            log.Info("start FormatEmailBody");
            var emailBody = EmailBodyHeader2;
            log.Info("[FormatEmailBody 2] " + emailBody);

            if (quotes.Count == 0)
            {
                return string.Empty;
            }

            long row = 1;
            foreach(Quote quote in quotes)
            {
                //                emailBody += string.Format(EmailBodyTableLineItem, quote.WorkOrder, quote.WorkTicket, quote.CustomerName, quote.SiteName, quote.DescriptionOfWork, quote.TicketNote);
                if (row == 1)
                {
                    emailBody += String.Format(EmailBodyTableItemBegin, "padding-top: 10px", quote.WorkOrder);
                }
                else
                {
                    if ((row % 2) == 0)
                    {
                        emailBody += String.Format(EmailBodyTableItemBegin, "background: lightgray", quote.WorkOrder);
                    }
                    else
                    {
                        emailBody += String.Format(EmailBodyTableItemBegin, "background: white", quote.WorkOrder);
                    }
                }
                log.Info("[FormatEmailBody 3] " + emailBody);

                emailBody += (quote.WorkTicket.Length > 0) ? string.Format(EmailBodyTableInnerItem2, "Work Ticket:", quote.WorkTicket) : string.Empty;
                emailBody += (quote.SiteName.Length > 0) ? string.Format(EmailBodyTableInnerItem2, "Site Name:", quote.SiteName) : string.Empty;
                emailBody += (quote.TicketNote.Length > 0) ? string.Format(EmailBodyTableInnerItem2, "Ticket Note:", quote.TicketNote) : string.Empty;
                emailBody += (quote.DescriptionOfWork.Length > 0) ? string.Format(EmailBodyTableInnerItem2, "Work Ticket:", quote.DescriptionOfWork) : string.Empty;

                emailBody += EmailBodyTablItemEnd;

                row++;
            }

            emailBody += EmailBodyFooter2;

            log.Info(emailBody);

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
