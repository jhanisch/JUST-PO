using JUST.Shared.DatabaseRepository;
using Moq;
using NUnit.Framework;
using System.IO;
using System.Collections.Generic;
using System.Net.Mail;

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
    }
}
