using JUST.Shared.Classes;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using NReco;
using MailKit;
using System.Text;
using System.Threading.Tasks;
using MimeKit;
using MailKit.Security;

namespace JUST.Shared.Utilities
{
    public static class EmployeeLookup
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public static Employee FindEmployeeFromAllEmployees(List<Employee> EmployeeEmailAddresses, string employee)
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
    }

    public static class Utils
    { 
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public static string createPdf(string HTML, string attachFilenamePrefix)
        {
            var outputPath = AppDomain.CurrentDomain.BaseDirectory + "DataFiles\\";
            System.IO.Directory.CreateDirectory(outputPath);
            var outputFile = outputPath + attachFilenamePrefix + "_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".pdf";

            var htmlToPdf = new NReco.PdfGenerator.HtmlToPdfConverter();
            htmlToPdf.GeneratePdf(HTML, null, outputFile);

            return outputFile;
        }

        public static bool sendEmail(Config config, ArrayList toEmailAddresses, string subject, string emailBody)
        {
            bool result = true;
            if (toEmailAddresses.Count == 0)
            {
                return false;
            }

            try
            {
                using (MailMessage mail = new MailMessage())
                {
                    mail.From = new MailAddress(config.FromEmailAddress, "New Job Notification");
                    foreach (string emailAddress in toEmailAddresses)
                    {
                        mail.To.Add(emailAddress);
                    }
                    mail.Subject = subject;
                    mail.Body = emailBody;
                    mail.IsBodyHtml = true;

                    using (SmtpClient smtp = new SmtpClient(config.FromEmailSMTP, config.FromEmailPort.Value))
                    {
                        smtp.Credentials = new NetworkCredential(config.FromEmailAddress, config.FromEmailPassword);
                        smtp.EnableSsl = true;
                        smtp.Send(mail);
                    }
                }
            }
            catch (Exception ex)
            {
                result = false;
                log.Debug("[SendEmail] - " + ex.Message);
            }

            return result;
        }

        public static bool sendEmail2(Config config, ArrayList toEmailAddresses, string subject, string emailBody, List<string> attachments)
        {
            bool result = true;
            if (toEmailAddresses.Count == 0)
            {
                return false;
            }

            try
            {
                var message = new MimeMessage();
                message.From.Add(new MailboxAddress("Notifications", config.FromEmailAddress));
                foreach (string emailAddress in toEmailAddresses)
                {
                    message.To.Add(new MailboxAddress(emailAddress, emailAddress));
                }
                message.Subject = subject;

                var builder = new BodyBuilder();
                builder.HtmlBody = emailBody;

/*                if (attachAsPdf)
                {
                    var outputPath = AppDomain.CurrentDomain.BaseDirectory + "DataFiles\\";
                    System.IO.Directory.CreateDirectory(outputPath);
                    var outputFile = outputPath + attachFilenamePrefix + "_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".pdf";

                    var htmlToPdf = new NReco.PdfGenerator.HtmlToPdfConverter();
                    htmlToPdf.GeneratePdf(emailBody, null, outputFile);

                    // Render any HTML fragment or document to HTML
                    builder.Attachments.Add(outputFile);
                }*/
                foreach(string attachment in attachments)
                {
                    builder.Attachments.Add(attachment);
                }

                message.Body = builder.ToMessageBody();

                using (MailKit.Net.Smtp.SmtpClient smtp = new MailKit.Net.Smtp.SmtpClient())
                {
                    smtp.Connect(config.FromEmailSMTP, config.FromEmailPort.Value, SecureSocketOptions.StartTls);
                    smtp.AuthenticationMechanisms.Remove("XOAUTH");
                    smtp.Authenticate(config.FromEmailAddress, config.FromEmailPassword);
                    smtp.Send(message);
                    smtp.Disconnect(true);
                }

            }
            catch (Exception ex)
            {
                result = false;
                log.Debug("[SendEmail2] - " + ex.Message);
            }

            return result;
        }

    }
}
