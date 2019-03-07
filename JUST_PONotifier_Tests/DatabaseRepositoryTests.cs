using JUST.Shared.DatabaseRepository;
using Moq;
using NUnit.Framework;
using System.IO;
using System.Collections.Generic;
using System.Net.Mail;
using JUST.Shared.Classes;

namespace JUST_PONotifier_Tests
{
    [TestFixture()]
    public class DatabaseRepositoryTests
    {
        public IDatabaseRepository MockDatabaseQueries;
        Mock<IDatabaseRepository> mockDatabaseQueries;

        [SetUp]
        public void TestSetup()
        {
            mockDatabaseQueries = new Mock<IDatabaseRepository>();
            Stream emptyStream = new MemoryStream();
            List<Attachment> mockPoAttachments = new List<Attachment>
            {
                new Attachment(emptyStream, "1")
            };

            mockDatabaseQueries.Setup(mdq => mdq.GetAttachmentsForPO("1")).Returns(mockPoAttachments);

            this.MockDatabaseQueries = mockDatabaseQueries.Object;
        }

        [TestCase("1", 1)]
        public void DatabaseQueries_GetAttachmentsForPo(string attachmentId, long expectedCount)
        {
            var result = MockDatabaseQueries.GetAttachmentsForPO(attachmentId);
            Assert.IsInstanceOf(typeof(List<Attachment>), result);
            Assert.AreEqual(result.Count, expectedCount);
        }

        [TestCase]
        public void DatabaseQueries_VerifyPOQuery()
        {
            string expectedPOQuery = "Select icpo.buyer, icpo.ponum, icpo.user_1, icpo.user_2, icpo.user_3, icpo.user_4, icpo.user_5, icpo.defaultjobnum, vendor.name as vendorName, icpo.user_6, icpo.defaultworkorder, icpo.attachid from icpo inner join vendor on vendor.vennum = icpo.vennum where icpo.user_3 is not null and icpo.user_5 = 0 order by icpo.ponum asc";

            DatabaseRepository testClass = new DatabaseRepository();
            Assert.AreEqual(testClass.POQuery, expectedPOQuery);
        }
    }
}
