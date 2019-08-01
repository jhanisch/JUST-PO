using JUST.Shared.Classes;
using NUnit.Framework;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Net.Mail;

namespace JUST.Shared.Tests
{
    [TestFixture()]
    public class ClassesTests
    {
        public List<Attachment> emptyList;
        private ArrayList ValidModes = new ArrayList(new string[] { "debug", "live", "monitor"});
        private DateTime dateOne = DateTime.Now.AddDays(-1);
        private DateTime dateTwo = DateTime.Now.AddDays(-2);

        [SetUp]
        public void TestSetup()
        {
            emptyList = new List<Attachment>();
        }

        #region Classes

        #region PurchaseOrder
        [Test]
        public void Classes_PurchaseOrderExists()
        {
            var newObject = new PurchaseOrder();
            Assert.AreEqual(String.Empty, newObject.PurchaseOrderNumber);
            Assert.AreEqual(String.Empty, newObject.ReceivedBy);
            Assert.AreEqual(String.Empty, newObject.Bin);
            Assert.AreEqual(String.Empty, newObject.ReceivedOnDate);
            Assert.AreEqual(String.Empty, newObject.Buyer);
            Assert.AreEqual(String.Empty, newObject.Vendor);
            Assert.AreEqual(String.Empty, newObject.Notes);
            Assert.AreEqual(String.Empty, newObject.WorkOrderNumber);
            Assert.AreEqual(String.Empty, newObject.BuyerEmployee);
            Assert.AreEqual(String.Empty, newObject.ProjectManagerEmployee);
            Assert.AreEqual(0, newObject.Attachments.Count);
        }

        [TestCase("12345", "Enzo Ferrari", "Bin 123", "01/01/2001", "aaa", "bbb", "asdfasdfasdfasdf", "11111111111111111", "WWWwww\\@!~!{}[]", "!@#$!$#%^&&^*", "jobNumber")]
        public void Classes_PurchaseOrderInitializes(
            string poNum, 
            string receivedBy, 
            string bin, 
            string receivedOnDate,
            string buyer,
            string vendor,
            string notes,
            string workOrderNumber,
            string buyerEmployee,
            string projectManagerEmployee,
            string jobNumber
            )
        {

            var newObject = new PurchaseOrder(
                poNum, 
                receivedBy, 
                bin, 
                receivedOnDate,
                buyer,
                vendor,
                notes,
                workOrderNumber,
                buyerEmployee,
                projectManagerEmployee,
                jobNumber,
                emptyList
                );

            Assert.AreEqual(poNum, newObject.PurchaseOrderNumber);
            Assert.AreEqual(receivedBy, newObject.ReceivedBy);
            Assert.AreEqual(bin, newObject.Bin);
            Assert.AreEqual(receivedOnDate, newObject.ReceivedOnDate);
            Assert.AreEqual(buyer, newObject.Buyer);
            Assert.AreEqual(vendor, newObject.Vendor);
            Assert.AreEqual(notes, newObject.Notes);
            Assert.AreEqual(workOrderNumber, newObject.WorkOrderNumber);
            Assert.AreEqual(buyerEmployee, newObject.BuyerEmployee);
            Assert.AreEqual(projectManagerEmployee, newObject.ProjectManagerEmployee);
            Assert.AreEqual(jobNumber, newObject.JobNumber);
            Assert.AreEqual(emptyList.Count, newObject.Attachments.Count);
        }
        #endregion

        #region Employee Lookup
        [Test]
        public void Classes_EmployeeExists()
        {
            var newObject = new Employee();
            Assert.AreEqual(String.Empty, newObject.EmployeeId);
            Assert.AreEqual(String.Empty, newObject.Name);
            Assert.AreEqual(String.Empty, newObject.EmailAddress);
        }

        [TestCase("12345", "Enzo Ferrari", "enzo@ferrari.com")]
        [TestCase("12345", null, "enzo@ferrari.com")]
        [TestCase("12345", "Enzo Ferrari", null)]
        public void Classes_EmployeeInitializes(
            string employeeId,
            string name,
            string emailAddress)
        {
            var newObject = new Employee(employeeId, name, emailAddress);

            Assert.AreEqual(employeeId, newObject.EmployeeId);
            Assert.AreEqual(name, newObject.Name);
            Assert.AreEqual(emailAddress, newObject.EmailAddress);
        }
        #endregion

        #region Config
        [TestCase]
        public void Classes_ConfigInitializes()
        {
            var newObject = new Config();

            Assert.AreEqual(newObject.FromEmailAddress, string.Empty);
            Assert.AreEqual(newObject.Uid, string.Empty);
            Assert.AreEqual(newObject.Pwd, string.Empty);
            Assert.AreEqual(newObject.POAttachmentBasePath, string.Empty);
            Assert.AreEqual(newObject.MonitorEmailAddresses.Count, 0);
            Assert.AreEqual(newObject.Mode, string.Empty);
            Assert.AreEqual(newObject.FromEmailSMTP, string.Empty);
            Assert.AreEqual(newObject.FromEmailPort, 25);
            Assert.AreEqual(newObject.FromEmailPassword, string.Empty);
            Assert.AreEqual(newObject.FromEmailAddress, string.Empty);
            Assert.AreEqual(newObject.ModeRequired, true);
        }
        
        //UID Validation
        [TestCase("", "pwd", "fromEmailAddress", "fromEmailPassword", "fromEmailSMTP", "0", "debug", "monitorEmailAddress", "C:\\", "User ID (Uid) is Required")]
        [TestCase("", "pwd", "fromEmailAddress", "fromEmailPassword", "fromEmailSMTP", "0", "Debug", "monitorEmailAddress", "C:\\", "User ID (Uid) is Required")]
        [TestCase("", "pwd", "fromEmailAddress", "fromEmailPassword", "fromEmailSMTP", "0", "DEBUG", "monitorEmailAddress", "C:\\", "User ID (Uid) is Required")]
        [TestCase(null, "pwd", "fromEmailAddress", "fromEmailPassword", "fromEmailSMTP", "0", "debug", "monitorEmailAddress", "C:\\", "User ID (Uid) is Required")]
        [TestCase("", "pwd", "fromEmailAddress", "fromEmailPassword", "fromEmailSMTP", "0", "debug", "", "C:\\", "User ID (Uid) is Required")]
        [TestCase("", "pwd", "fromEmailAddress", "fromEmailPassword", "fromEmailSMTP", "0", "live", "", "C:\\", "User ID (Uid) is Required")]
        [TestCase("", "pwd", "fromEmailAddress", "fromEmailPassword", "fromEmailSMTP", "0", "monitor", "montorEmailAddress", "C:\\", "User ID (Uid) is Required")]
        [TestCase("", "pwd", "fromEmailAddress", "fromEmailPassword", "fromEmailSMTP", "0", "monitor", "", "C:\\", "User ID (Uid) is RequiredMonitor Email Address is Required in monitor mode")]
        [TestCase("", "pwd", "fromEmailAddress", "fromEmailPassword", "fromEmailSMTP", "0", "monitor", "a,b,c", "C:\\", "User ID (Uid) is Required")]
        [TestCase("", "pwd", "fromEmailAddress", "fromEmailPassword", "fromEmailSMTP", "0", "monitor", "a;b;c", "C:\\", "User ID (Uid) is Required")]
        //Password validation
        [TestCase("uid", "", "fromEmailAddress", "fromEmailPassword", "fromEmailSMTP", "0", "debug", "monitorEmailAddress", "C:\\", "Password (Pwd) is Required")]
        [TestCase("uid", "", "fromEmailAddress", "fromEmailPassword", "fromEmailSMTP", "0", "debug", "", "C:\\", "Password (Pwd) is Required")]
        [TestCase("uid", "pwd", "", "fromEmailPassword", "fromEmailSMTP", "0", "debug", "monitorEmailAddress", "C:\\", "From Email Address (FromEmailAddress) is Required")]
        //FromEmailAddress
        [TestCase("uid", "pwd", "", "fromEmailPassword", "fromEmailSMTP", "0", "debug", "monitorEmailAddress", "C:\\", "From Email Address (FromEmailAddress) is Required")]
        [TestCase("uid", "pwd", null, "fromEmailPassword", "fromEmailSMTP", "0", "debug", "monitorEmailAddress", "C:\\", "From Email Address (FromEmailAddress) is Required")]
        //FromEmailPassword
        [TestCase("uid", "pwd", "fromEmailAddress", "", "fromEmailSMTP", "0", "debug", "monitorEmailAddress", "C:\\", "From Email Password (FromEmailPassword) is Required")]
        [TestCase("uid", "pwd", "fromEmailAddress", null, "fromEmailSMTP", "0", "debug", "monitorEmailAddress", "C:\\", "From Email Password (FromEmailPassword) is Required")]
        //FromEmailSMTP
        [TestCase("uid", "pwd", "fromEmailAddress", "fromEmailPassword", "", "0", "debug", "monitorEmailAddress", "C:\\", "From Email SMTP (FromEmailSMTP) address is Required")]
        [TestCase("uid", "pwd", "fromEmailAddress", "fromEmailPassword", null, "0", "debug", "monitorEmailAddress", "C:\\", "From Email SMTP (FromEmailSMTP) address is Required")]
        //FromEmailPort
        [TestCase("uid", "pwd", "fromEmailAddress", "fromEmailPassword", "fromEmailSMTP", "", "debug", "monitorEmailAddress", "C:\\", "From Email Port (FromEmailPort) is Required")]
        [TestCase("uid", "pwd", "fromEmailAddress", "fromEmailPassword", "fromEmailSMTP", null, "debug", "monitorEmailAddress", "C:\\", "From Email Port (FromEmailPort) is Required")]
        [TestCase("uid", "pwd", "fromEmailAddress", "fromEmailPassword", "fromEmailSMTP", "-1", "debug", "monitorEmailAddress", "C:\\", "From Email Port (FromEmailPort) must be a positive value")]
        //Mode
        [TestCase("uid", "pwd", "fromEmailAddress", "fromEmailPassword", "fromEmailSMTP", "0", "monitor", "", "C:\\", "Monitor Email Address is Required in monitor mode")]
        [TestCase("uid", "pwd", "fromEmailAddress", "fromEmailPassword", "fromEmailSMTP", "0", "Monitor", "", "C:\\", "Monitor Email Address is Required in monitor mode")]
        [TestCase("uid", "pwd", "fromEmailAddress", "fromEmailPassword", "fromEmailSMTP", "0", "MONITOR", "", "C:\\", "Monitor Email Address is Required in monitor mode")]
        [TestCase("uid", "pwd", "fromEmailAddress", "fromEmailPassword", "fromEmailSMTP", "0", "", "", "C:\\", "Mode is Required.")]
        [TestCase("uid", "pwd", "fromEmailAddress", "fromEmailPassword", "fromEmailSMTP", "0", null, "", "C:\\", "Mode is Required.")]
        // MonitorEmailAddress
        [TestCase("uid", "pwd", "fromEmailAddress", "fromEmailPassword", "fromEmailSMTP", "0", "monitor", null, "C:\\", "Monitor Email Address is Required in monitor mode")]
        public void Classes_Config_ThrowsExceptions(
            string uid, 
            string pwd, 
            string fromEmailAddress, 
            string fromEmailPassword, 
            string fromEmailSMTP, 
            string fromEmailPort, 
            string mode,
            string monitorEmailAddress,
            string poAttachmentBasePath,
            string expectedExceptionMessage)
        {
            var newObject = new Config();

            ConfigurationManager.AppSettings["Uid"] = uid;
            ConfigurationManager.AppSettings["Pwd"] = pwd;
            ConfigurationManager.AppSettings["FromEmailAddress"] = fromEmailAddress;
            ConfigurationManager.AppSettings["FromEmailPassword"] = fromEmailPassword; 
            ConfigurationManager.AppSettings["FromEmailSMTP"] = fromEmailSMTP;
            ConfigurationManager.AppSettings["FromEmailPort"] = fromEmailPort;
            ConfigurationManager.AppSettings["Mode"] = mode;
            ConfigurationManager.AppSettings["MonitorEmailAddress"] = monitorEmailAddress;
            ConfigurationManager.AppSettings["POAttachmentBasePath"] = poAttachmentBasePath;

            var result = Assert.Throws<Exception>(() => newObject.getConfiguration(ValidModes));
            Assert.That(result.Message, Is.EqualTo(expectedExceptionMessage));

            if (uid != null)
            {
                Assert.AreEqual(newObject.Uid, uid);
            }
            else
            {
                Assert.IsNull(newObject.Uid);
            }

            if (pwd != null)
            {
                Assert.AreEqual(newObject.Pwd, pwd);
            }
            else
            {
                Assert.IsNull(newObject.Pwd);
            }

            if (fromEmailAddress != null)
            {
                Assert.AreEqual(newObject.FromEmailAddress, fromEmailAddress);
            }
            else
            {
                Assert.IsNull(newObject.FromEmailAddress);
            }

            if (fromEmailPassword != null)
            {
                Assert.AreEqual(newObject.FromEmailPassword, fromEmailPassword);
            }
            else
            {
                Assert.IsNull(newObject.FromEmailPassword);
            }

            if (fromEmailSMTP != null)
            {
                Assert.AreEqual(newObject.FromEmailSMTP, fromEmailSMTP);
            }
            else
            {
                Assert.IsNull(newObject.FromEmailSMTP);
            }

            if (!String.IsNullOrEmpty(fromEmailPort))
            {
                Assert.AreEqual(newObject.FromEmailPort, Convert.ToInt32(fromEmailPort));
            }
            else
            {
                Assert.AreEqual(newObject.FromEmailPort, 25);  // email port defaults to 25 if no value is found
            }

            /*  need to test monitor email address splitting
            if (monitorEmailAddress != null)
            {
                Assert.AreEqual(newObject.MonitorEmailAddresses, monitorEmailAddress);
            }
            else
            {
                Assert.IsNull(newObject.MonitorEmailAddresses);
            }
            */

            if (poAttachmentBasePath != null)
            {
                Assert.AreEqual(newObject.POAttachmentBasePath, poAttachmentBasePath);
            }
            else
            {
                Assert.IsNull(newObject.POAttachmentBasePath);
            }

            if (mode != null)
            {
                Assert.AreEqual(newObject.Mode, mode.ToLower());
            }
            else
            {
                Assert.IsEmpty(newObject.Mode);
            }
        }

        [TestCase(true)]
        [TestCase(false)]
        public void Classes_Config_SetsModeRequired(bool required)
        {
            var newObject = new Config(required);

            Assert.AreEqual(newObject.ModeRequired, required);
        }

        [TestCase("uid", "pwd", "fromEmailAddress", "fromEmailPassword", "fromEmailSMTP", "0", "kjhkjh", null, "C:\\")]
        public void Classes_Config_DoesNotRequireMode(
            string uid,
            string pwd,
            string fromEmailAddress,
            string fromEmailPassword,
            string fromEmailSMTP,
            string fromEmailPort,
            string mode,
            string monitorEmailAddress,
            string poAttachmentBasePath)
        {
            ConfigurationManager.AppSettings["Uid"] = uid;
            ConfigurationManager.AppSettings["Pwd"] = pwd;
            ConfigurationManager.AppSettings["FromEmailAddress"] = fromEmailAddress;
            ConfigurationManager.AppSettings["FromEmailPassword"] = fromEmailPassword;
            ConfigurationManager.AppSettings["FromEmailSMTP"] = fromEmailSMTP;
            ConfigurationManager.AppSettings["FromEmailPort"] = fromEmailPort;
            ConfigurationManager.AppSettings["Mode"] = mode;
            ConfigurationManager.AppSettings["MonitorEmailAddress"] = monitorEmailAddress;
            ConfigurationManager.AppSettings["POAttachmentBasePath"] = poAttachmentBasePath;

            var newObject = new Config(false);

            var result = newObject.getConfiguration(ValidModes);

            Assert.IsTrue(result);
            Assert.AreEqual(string.Empty, newObject.Mode);
            Assert.IsFalse(newObject.ModeRequired);
        }

        [TestCase("uid", "pwd", "fromEmailAddress", "fromEmailPassword", "fromEmailSMTP", "0", "monitor", "a;b", "C:\\")]
        [TestCase("uid", "pwd", "fromEmailAddress", "fromEmailPassword", "fromEmailSMTP", "0", "monitor", "a,b", "C:\\")]
        public void Classes_Config_ParsesEmailAddresses(
            string uid,
            string pwd,
            string fromEmailAddress,
            string fromEmailPassword,
            string fromEmailSMTP,
            string fromEmailPort,
            string mode,
            string monitorEmailAddress,
            string poAttachmentBasePath)
        {
            ConfigurationManager.AppSettings["Uid"] = uid;
            ConfigurationManager.AppSettings["Pwd"] = pwd;
            ConfigurationManager.AppSettings["FromEmailAddress"] = fromEmailAddress;
            ConfigurationManager.AppSettings["FromEmailPassword"] = fromEmailPassword;
            ConfigurationManager.AppSettings["FromEmailSMTP"] = fromEmailSMTP;
            ConfigurationManager.AppSettings["FromEmailPort"] = fromEmailPort;
            ConfigurationManager.AppSettings["Mode"] = mode;
            ConfigurationManager.AppSettings["MonitorEmailAddress"] = monitorEmailAddress;
            ConfigurationManager.AppSettings["POAttachmentBasePath"] = poAttachmentBasePath;

            var newObject = new Config(false);

            var result = newObject.getConfiguration(ValidModes);

            Assert.IsTrue(result);
            Assert.AreEqual(2, newObject.MonitorEmailAddresses.Count);
        }

        #endregion

        #region Quote Needed
        [Test]
        public void Classes_QuoteNeeded()
        {
            var newObject = new Quote();
            Assert.AreEqual(String.Empty, newObject.WorkOrder);
            Assert.AreEqual(String.Empty, newObject.WorkTicket);
            Assert.AreEqual(String.Empty, newObject.CustomerName);
            Assert.AreEqual(String.Empty, newObject.SiteName);
            Assert.AreEqual(String.Empty, newObject.DescriptionOfWork);
            Assert.AreEqual(String.Empty, newObject.TicketNote);
            Assert.AreEqual(String.Empty, newObject.ServiceTech);
            Assert.AreEqual(String.Empty, newObject.Manufacturer);
            Assert.AreEqual(String.Empty, newObject.Model);
            Assert.AreEqual(String.Empty, newObject.SerialNumber);
        }

        [TestCase("workOrder", "workTicket", "customerName", "siteName", "descriptionOfWork", "ticketNote", "serviceperson", "make", "model", "serialnumber")]
        public void Classes_QuoteNeeded(string workOrder, string workTicket, string customerName, string siteName, string descriptionOfWork, string ticketNote, string servicePerson, string manufacturer, string model, string serialNumber)
        {
            var newObject = new Quote(workOrder, workTicket, customerName, siteName, descriptionOfWork, ticketNote, servicePerson, manufacturer, model, serialNumber, string.Empty);

            Assert.AreEqual(workOrder, newObject.WorkOrder);
            Assert.AreEqual(workTicket, newObject.WorkTicket);
            Assert.AreEqual(customerName, newObject.CustomerName);
            Assert.AreEqual(siteName, newObject.SiteName);
            Assert.AreEqual(descriptionOfWork, newObject.DescriptionOfWork);
            Assert.AreEqual(ticketNote, newObject.TicketNote);
            Assert.AreEqual(servicePerson, newObject.ServiceTech);
            Assert.AreEqual(manufacturer, newObject.Manufacturer);
            Assert.AreEqual(model, newObject.Model);
            Assert.AreEqual(serialNumber, newObject.SerialNumber);

        }
        #endregion

        #region Aged Receivables
        [Test]
        public void Classes_AgedReceivables()
        {
            var newObject = new AgedReceivable();
            Assert.AreEqual(String.Empty, newObject.CustomerNumber);
            Assert.AreEqual(String.Empty, newObject.CustomerName);
            Assert.AreEqual(String.Empty, newObject.InvoiceNumber);
            Assert.AreEqual(String.Empty, newObject.JobNumber);
            Assert.AreEqual(String.Empty, newObject.WorkOrderNumber);
            Assert.AreEqual(0.00M, newObject.AmountDue);
            Assert.AreEqual(DateTime.Now.ToShortDateString(), newObject.InvoiceDate.ToShortDateString());
            Assert.AreEqual(String.Empty, newObject.Memo);
            Assert.AreEqual(false, newObject.Notified);
        }

        [TestCase("customerNumber", "customerName", "invoiceNumber", "jobNumber", "workOrderNumber", "memo")]
        public void Classes_AgedReceivables(string customerNumber, string customerName, string invoiceNumber, string jobNumber, string workOrderNumber, string memo)
        {
            var newObject = new AgedReceivable(customerNumber, customerName, invoiceNumber, dateOne, dateTwo, jobNumber, workOrderNumber, 100.00M, memo);

            Assert.AreEqual(customerNumber, newObject.CustomerNumber);
            Assert.AreEqual(customerName, newObject.CustomerName);
            Assert.AreEqual(invoiceNumber, newObject.InvoiceNumber);
            Assert.AreEqual(dateOne, newObject.InvoiceDate);
            Assert.AreEqual(dateTwo, newObject.DueDate);
            Assert.AreEqual(jobNumber, newObject.JobNumber);
            Assert.AreEqual(workOrderNumber, newObject.WorkOrderNumber);
            Assert.AreEqual(100.00M, newObject.AmountDue);
            Assert.AreEqual(memo, newObject.Memo);
            Assert.AreEqual(2, newObject.DaysOverdue);
            Assert.AreEqual(false, newObject.Notified);
        }

        [TestCase("customerNumber", "customerName", "invoiceNumber", "jobNumber", "workOrderNumber", "memo")]
        public void Classes_AgedReceivables_MinDate(string customerNumber, string customerName, string invoiceNumber, string jobNumber, string workOrderNumber, string memo)
        {
            var newObject = new AgedReceivable(customerNumber, customerName, invoiceNumber, dateOne, DateTime.MinValue, jobNumber, workOrderNumber, 100.00M, memo);

            Assert.AreEqual(-1, newObject.DaysOverdue);
        }


        #endregion

        #endregion
    }
}
