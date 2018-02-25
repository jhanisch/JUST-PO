using System;
using System.Collections.Generic;

namespace JUST.PONotifier.Classes
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
            PurchaseOrderItems = new List<PurchaseOrderItem>();
        }

        public JobInformation(string projectManagerName, string jobNumber, string jobName, string customerNumber)
        {
            ProjectManagerName = projectManagerName;
            JobNumber = jobNumber;
            JobName = jobName;
            CustomerNumber = customerNumber;
            PurchaseOrderItems = new List<PurchaseOrderItem>();
        }

        public string ProjectManagerName { get; set; }
        public string JobNumber { get; set; }
        public string JobName { get; set; }
        public string CustomerNumber { get; set; }
        public string CustomerName { get; set; }
        public IList<PurchaseOrderItem> PurchaseOrderItems { get; set; }
    }
}
