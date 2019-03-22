using JUST.Shared.Classes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

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
        public static bool sendEmail(Config config, string toEmailAddress, string subject, string emailBody)
        {
            bool result = true;
            if (toEmailAddress.Length == 0)
            {
                return false;
            }

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
                    }
                }
            }
            catch (Exception x)
            {
                result = false;
            }

            return result;
        }
    }
}
