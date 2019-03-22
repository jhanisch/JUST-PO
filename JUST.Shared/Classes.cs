using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Net.Mail;
using System.Text;

namespace JUST.Shared.Classes
{
    public class Config
    {
        private const string monitor = "monitor";

        public Config()
        {
            Uid = string.Empty;
            Pwd = string.Empty;
            FromEmailAddress = string.Empty;
            FromEmailPassword = string.Empty;
            FromEmailSMTP = string.Empty;
            FromEmailPort = 25;
            Mode = string.Empty;
            MonitorEmailAddresses = new string[0];
            POAttachmentBasePath = string.Empty;
        }

        public string Uid { get; private set; }
        public string Pwd { get; private set; }
        public string FromEmailAddress { get; private set; }
        public string FromEmailPassword { get; private set; }
        public string FromEmailSMTP { get; private set; }
        public int? FromEmailPort { get; private set; }
        public string Mode { get; private set; }
        public string[] MonitorEmailAddresses;
        public string POAttachmentBasePath { get; private set; }

        public void getConfiguration(ArrayList ValidModes, Boolean AttachmentBasePathRequired = true)
        {
            string fromEmailPortString, modeString;

            Uid = ConfigurationManager.AppSettings["Uid"];
            Pwd = ConfigurationManager.AppSettings["Pwd"];
            FromEmailAddress = ConfigurationManager.AppSettings["FromEmailAddress"];
            FromEmailPassword = ConfigurationManager.AppSettings["FromEmailPassword"];
            FromEmailSMTP = ConfigurationManager.AppSettings["FromEmailSMTP"];
            fromEmailPortString = ConfigurationManager.AppSettings["FromEmailPort"];

            modeString = ConfigurationManager.AppSettings["Mode"];
            var MonitorEmailAddressList = ConfigurationManager.AppSettings["MonitorEmailAddress"];
            if (MonitorEmailAddressList != null && MonitorEmailAddressList.Length > 0)
            {
                char[] delimiterChars = { ';', ',' };
                MonitorEmailAddresses = MonitorEmailAddressList.Split(delimiterChars);
            }

            POAttachmentBasePath = ConfigurationManager.AppSettings["POAttachmentBasePath"];

            #region Validate Configuration Data
            var errorMessage = new StringBuilder();
            if (String.IsNullOrEmpty(Uid))
            {
                errorMessage.Append("User ID (Uid) is Required");
            }

            if (String.IsNullOrEmpty(Pwd))
            {
                errorMessage.Append("Password (Pwd) is Required");
            }

            if (String.IsNullOrEmpty(FromEmailAddress))
            {
                errorMessage.Append("From Email Address (FromEmailAddress) is Required");
            }

            if (String.IsNullOrEmpty(FromEmailPassword))
            {
                errorMessage.Append("From Email Password (FromEmailPassword) is Required");
            }

            if (String.IsNullOrEmpty(FromEmailSMTP))
            {
                errorMessage.Append("From Email SMTP (FromEmailSMTP) address is Required");
            }

            if (String.IsNullOrEmpty(fromEmailPortString))
            {
                errorMessage.Append("From Email Port (FromEmailPort) is Required");
            }
            else
            {
                FromEmailPort = Convert.ToInt16(fromEmailPortString);

                if (FromEmailPort < 0)
                {
                    errorMessage.Append("From Email Port (FromEmailPort) must be a positive value");
                }
            }

            if (String.IsNullOrEmpty(modeString))
            {
                errorMessage.Append("Mode is Required.");
            }
            else
            {
                Mode = modeString.ToLower();
                if (!ValidModes.Contains(Mode))
                {
                    errorMessage.Append(String.Format("'{0}' is not a valid Mode.  Valid modes are 'debug', 'live' and 'monitor'", Mode));
                }
            }


            if (Mode == monitor)
            {
                if (MonitorEmailAddresses == null || MonitorEmailAddresses.Length == 0)
                {
                    errorMessage.Append("Monitor Email Address is Required in monitor mode");
                }
            }

            if (AttachmentBasePathRequired && String.IsNullOrEmpty(POAttachmentBasePath))
            {
                errorMessage.Append("Root Path to Attachments (AttachmentBasePath) is Required");
            }

            if (errorMessage.Length > 0)
            {
                throw new Exception(errorMessage.ToString());
            }
            #endregion
        }
    }

    public class Employee
    {
        public Employee()
        {
            EmployeeId = string.Empty;
            Name = string.Empty;
            EmailAddress = string.Empty;
        }

        public Employee(string employeeId, string name, string emailAddress)
        {
            EmployeeId = employeeId;
            Name = name;
            EmailAddress = emailAddress;
        }

        public string EmployeeId { get; set; }
        public string Name { get; set; }
        public string EmailAddress { get; set; }
    }

    public class PurchaseOrder
    {
        public PurchaseOrder()
        {
            PurchaseOrderNumber = String.Empty;
            ReceivedBy = String.Empty;
            Bin = String.Empty;
            ReceivedOnDate = String.Empty;
            Buyer = String.Empty;
            Vendor = String.Empty;
            Notes = String.Empty;
            WorkOrderNumber = String.Empty;
            BuyerEmployee = String.Empty;
            ProjectManagerEmployee = String.Empty;
            JobNumber = String.Empty;
            Attachments = new List<Attachment>();
        }

        public PurchaseOrder(
                string purchaseOrderNumber,
                string receivedBy,
                string bin,
                string receivedOnDate,
                string buyer,
                string vendor,
                string notes,
                string workOrderNumber,
                string buyerEmployee,
                string projectManagerEmployee,
                string jobNumber,
                List<Attachment> attachments
            )
        {
            PurchaseOrderNumber = purchaseOrderNumber;
            ReceivedBy = receivedBy;
            Bin = bin;
            ReceivedOnDate = receivedOnDate;
            Buyer = buyer;
            Vendor = vendor;
            Notes = notes;
            WorkOrderNumber = workOrderNumber;
            BuyerEmployee = buyerEmployee;
            ProjectManagerEmployee = projectManagerEmployee;
            JobNumber = jobNumber;

            Attachments = new List<Attachment>();
            Attachments.AddRange(attachments);
        }

        public string PurchaseOrderNumber { get; set; }
        public string ReceivedBy { get; set; }
        public string Bin { get; set; }
        public string ReceivedOnDate { get; set; }
        public string Buyer { get; set; }
        public string Vendor { get; set; }
        public string Notes { get; set; }
        public string WorkOrderNumber { get; set; }
        public string BuyerEmployee { get; set; }
        public string ProjectManagerEmployee { get; set; }
        public string JobNumber { get; set; }
        public List<Attachment> Attachments { get; set; }
    }

    public class PurchaseOrderItem
    {
        public PurchaseOrderItem()
        {
            ItemNumber = string.Empty;
            Description = string.Empty;
            Quantity = string.Empty;
            UnitPrice = string.Empty;
        }

        public PurchaseOrderItem (string itemNumber, string description, string quantity = "", string unitPrice = "")
        {
            ItemNumber = itemNumber;
            Description = description;
            Quantity = quantity;
            UnitPrice = unitPrice;
        }

        public string ItemNumber { get; set; }
        public string Description { get; set; }
        public string Quantity { get; set; }
        public string UnitPrice { get; set; }
        public long Received { get; set; }
    }

    public class JobInformation
    {
         public JobInformation()
        {
            ProjectManagerName = string.Empty;
            JobNumber = string.Empty;
            JobName = string.Empty;
            CustomerNumber = string.Empty;
            CustomerName = string.Empty;
            WorkOrderNumber = string.Empty;
            WorkOrderSite = string.Empty;
            PurchaseOrderItems = new List<PurchaseOrderItem>();
        }

        public JobInformation(string projectManagerName, string jobNumber, string jobName, string customerNumber, string workOrderNumber, string workOrderSite)
        {
            ProjectManagerName = projectManagerName;
            JobNumber = jobNumber;
            JobName = jobName;
            CustomerNumber = customerNumber;
            WorkOrderNumber = workOrderNumber;
            WorkOrderSite = workOrderSite;
            PurchaseOrderItems = new List<PurchaseOrderItem>();
        }

        public string ProjectManagerName { get; set; }
        public string JobNumber { get; set; }
        public string JobName { get; set; }
        public string CustomerNumber { get; set; }
        public string CustomerName { get; set; }
        public string WorkOrderNumber { get; set; }
        public string WorkOrderSite { get; set; }
        public IList<PurchaseOrderItem> PurchaseOrderItems { get; set; }
    }

    public class Quote
    {
        public string WorkOrder { get; set; }
        public string WorkTicket { get; set; }
        public string CustomerName { get; set; }
        public string SiteName { get; set; }
        public string DescriptionOfWork { get; set; }
        public string TicketNote { get; set; }
        public string ServicePerson { get; set; }
        public string Manufacturer { get; set; }
        public string Model { get; set; }
        public string SerialNumber { get; set; }

        public Quote()
        {
            WorkOrder = string.Empty;
            WorkTicket = string.Empty;
            CustomerName = string.Empty;
            SiteName = string.Empty;
            DescriptionOfWork = string.Empty;
            TicketNote = string.Empty;
            ServicePerson = string.Empty;
            Manufacturer = string.Empty;
            Model = string.Empty;
            SerialNumber = string.Empty;
        }

        public Quote(string workOrder, 
            string workTicket, 
            string customerName, 
            string siteName, 
            string descriptionOfWork, 
            string ticketNote,
            string servicePerson,
            string manufacturer,
            string model,
            string serialNumber)
        {
            WorkOrder = workOrder;
            WorkTicket = workTicket;
            CustomerName = customerName;
            SiteName = siteName;
            DescriptionOfWork = descriptionOfWork;
            TicketNote = ticketNote;
            ServicePerson = servicePerson;
            Manufacturer = manufacturer;
            Model = model;
            SerialNumber = serialNumber;
        }

    }
}
