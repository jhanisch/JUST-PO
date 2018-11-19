using NUnit.Framework;
using System;
using JUST_PONotifier.Classes;

namespace JUST_PONotifier_Tests
{
    [TestFixture()]
    public class Tests
    {
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
            Assert.AreEqual(String.Empty, newObject.AttachmentId);
            Assert.AreEqual(String.Empty, newObject.BuyerEmployee);
            Assert.AreEqual(String.Empty, newObject.ProjectManagerEmployee);
        }

        [TestCase("12345", "Enzo Ferrari", "Bin 123", "01/01/2001", "aaa", "bbb", "asdfasdfasdfasdf", "11111111111111111", "WWWwww\\@!~!{}[]", "!@#$!$#%^&&^*", "ZXCV")]
        public void Classes_PurchaseOrderInitializes(
            string poNum, 
            string receivedBy, 
            string bin, 
            string receivedOnDate,
            string buyer,
            string vendor,
            string notes,
            string workOrderNumber,
            string attachmentId,
            string buyerEmployee,
            string projectManagerEmployee
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
                attachmentId,
                buyerEmployee,
                projectManagerEmployee);

            Assert.AreEqual(poNum, newObject.PurchaseOrderNumber);
            Assert.AreEqual(receivedBy, newObject.ReceivedBy);
            Assert.AreEqual(bin, newObject.Bin);
            Assert.AreEqual(receivedOnDate, newObject.ReceivedOnDate);
            Assert.AreEqual(buyer, newObject.Buyer);
            Assert.AreEqual(vendor, newObject.Vendor);
            Assert.AreEqual(notes, newObject.Notes);
            Assert.AreEqual(workOrderNumber, newObject.WorkOrderNumber);
            Assert.AreEqual(attachmentId, newObject.AttachmentId);
            Assert.AreEqual(buyerEmployee, newObject.BuyerEmployee);
            Assert.AreEqual(projectManagerEmployee, newObject.ProjectManagerEmployee);
        }
        #endregion
    }
}
