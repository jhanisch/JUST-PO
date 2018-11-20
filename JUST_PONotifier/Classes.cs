using System;
using System.Net.Mail;
using System.Collections.Generic;

namespace JUST_PONotifier.Classes
{

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
}
