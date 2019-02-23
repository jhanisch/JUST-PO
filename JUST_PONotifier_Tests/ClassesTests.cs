using JUST.Shared.Classes;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Net.Mail;

namespace JUST_PONotifier_Tests
{
    [TestFixture()]
    public class ClassesTests
    {
        public List<Attachment> emptyList;

        [SetUp]
        public void TestSetup()
        {
            emptyList = new List<Attachment>();
        }

        #region Classes
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
        }
}
