using JUST.Shared.Classes;
using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Linq;
using System.Net.Mail;

namespace JUST.Shared.DatabaseRepository
{
    public class DatabaseRepository : IDatabaseRepository
    {
        private OdbcConnection dbConnection;
        private log4net.ILog log;
        private string POAttachmentsRootPath;
        private List<Employee> Employees = new List<Employee>();
        private string POQueryString = "Select icpo.buyer, icpo.ponum, icpo.user_1, icpo.user_2, icpo.user_3, icpo.user_4, icpo.user_5, icpo.defaultjobnum, vendor.name as vendorName, icpo.user_6, icpo.defaultworkorder, icpo.attachid from icpo inner join vendor on vendor.vennum = icpo.vennum where icpo.user_3 is not null and icpo.user_5 = 0 order by icpo.ponum asc";
        private string QuotesNeededQueryString = "select dpwoassign.workorder workorder, dpwoassign.ticketnum ticketnum, customer.name customername, dpsite.sitenum sitenum, dpsite.name sitename, dporder.des workorderdescription, dpwoassign.user_17 ticketnote, dpwoassign.user_2 model, dpwoassign.user_3 serialnumber, dpwoassign.user_4 manufacturer, dpwoassign.person servicetech, dpwoassign.date workdate from dpwoassign, dporder, dpsite, customer where dpwoassign.user_15 = 1 and dpwoassign.user_18 = 0 and dpwoassign.ticketnum<> '' and dporder.workorder = dpwoassign.workorder and dpsite.sitenum = dpwoassign.sitenum and customer.cusnum = dpsite.cusnum";

        public string POQuery
        {
            get { return POQueryString; }
        }

        public string QuotesNeededQuery
        {
            get { return QuotesNeededQueryString; }
        }

        public DatabaseRepository()
        { }

        public DatabaseRepository(OdbcConnection cn, log4net.ILog x, string poAttachmentsRootPath)
        {
            dbConnection = cn;
            log = x;
            POAttachmentsRootPath = poAttachmentsRootPath;
            Employees = GetEmployees();
        }

        public List<Attachment> GetAttachmentsForPO(string attachid)
        {
            return GetAttachmentsForPO(dbConnection, attachid);
        }

        private List<Attachment> GetAttachmentsForPO(OdbcConnection cn, string attachid)
        {
            List<Attachment> poAttachments = new List<Attachment>();

            if (cn == null)
            {
                throw new Exception("GetAttachmentsForPO - No Database connection exists");
            }

            if (attachid.Trim().Length == 0)
            {
                log.Info("GetAttachmentsForPO, attachid is empty");
                return poAttachments;
            }

            var attachmentQuery = "Select path, displayname from icpoattachment where ownerkey = {0}";
            var attachmentCmd = new OdbcCommand(string.Format(attachmentQuery, attachid), cn);

            OdbcDataReader attachmentReader = attachmentCmd.ExecuteReader();

            while (attachmentReader.Read())
            {
                var path = attachmentReader.GetString(0);
                var displayName = attachmentReader.GetString(1);

                if (path.Trim().Length > 0)
                {
                    var fullPath = POAttachmentsRootPath + path;
                    poAttachments.Add(new Attachment(fullPath));
                }
            }

            attachmentReader.Close();

            return poAttachments;
        }

        public JobInformation GetEmailBodyInformation(string jobNum, string purchaseOrderNumber, string workOrderNumber)
        {
            return GetEmailBodyInformation(dbConnection, jobNum, purchaseOrderNumber, workOrderNumber);
        }

        private JobInformation GetEmailBodyInformation(OdbcConnection cn, string jobNum, string purchaseOrderNumber, string workOrderNumber)
        {
            // from the jcjob table (Company 0)
            // user_1 = job description
            // user_2 = sales person
            // user_3 = designer
            // user_4 = Project Manager
            // user_5 = SM
            // user_6 = Fitter
            // user_7 = Plumber
            // user_8 = Tech 1
            // user_9 = Tech 2

            // add to email:
            //  user_1 
            var jobQuery = "Select jcjob.user_4, jcjob.user_1, customer.name, customer.cusNum from jcjob inner join customer on customer.cusnum = jcjob.cusnum where jobnum = '{0}'";
            var jobCmd = new OdbcCommand(string.Format(jobQuery, jobNum), cn);
            var result = new JobInformation();
            result.JobNumber = jobNum;

            var jobReader = jobCmd.ExecuteReader();

            if (jobReader.Read())
            {
                result.ProjectManagerName = jobReader.GetString(0).ToLower();
                result.JobName = jobReader.GetString(1);
                result.CustomerName = jobReader.GetString(2);
                result.CustomerNumber = jobReader.GetString(3);
            }
            else
            {
                log.Info(string.Format("  [GetEmailBodyInformation] ERROR: Job Number (jobnum) {0} not found in jcjob", jobNum));
            }

            jobReader.Close();

            if (workOrderNumber.Trim().Length > 0)
            {
                var workOrderQuery = "Select dporder.workorder, dpsite.name from dporder inner join dpsite on dpsite.sitenum = dporder.sitenum where workorder = '{0}'";
                var workOrderCmd = new OdbcCommand(string.Format(workOrderQuery, workOrderNumber), cn);
                var workOrderResult = new JobInformation();

                OdbcDataReader workOrderReader = workOrderCmd.ExecuteReader();

                if (workOrderReader.Read())
                {
                    result.WorkOrderNumber = workOrderReader.GetString(0);
                    result.WorkOrderSite = workOrderReader.GetString(1);
                    log.Info(string.Format("  [GetEmailBodyInformation] INFO: Work Order Number {0} found, Site {1}", workOrderNumber, result.WorkOrderSite));
                }
                else
                {
                    log.Info(string.Format("  [GetEmailBodyInformation] INFO: Work Order Number (ponum) {0} not found in work order", workOrderNumber));
                }

                workOrderReader.Close();
            }

            var poIitemQuery = "Select icpoitem.itemnum, icpoitem.des, icpoitem.outstanding, icpoitem.unitprice from icpoitem where ponum = '{0}' order by icpoitem.itemnum asc";
            var poItemCmd = new OdbcCommand(string.Format(poIitemQuery, purchaseOrderNumber), cn);
            OdbcDataReader poItemReader = poItemCmd.ExecuteReader();

            while (poItemReader.Read())
            {
                result.PurchaseOrderItems.Add(new PurchaseOrderItem(poItemReader.GetString(0), poItemReader.GetString(1), poItemReader.GetString(2), poItemReader.GetString(3)));
            }

            return result;
        }

        public List<Employee> GetEmployees()
        {
            if (Employees.Count > 0)
            {
                return Employees;
            }

            if (dbConnection == null)
            {
                throw new Exception("GetEmployees - No Database Connection");
            }

            var employees = new List<Employee>();

            var query = "Select user_1, user_2, name from premployee where user_1 is not null";
            var cmd = new OdbcCommand(query, dbConnection);

            OdbcDataReader buyerReader = cmd.ExecuteReader();

            while (buyerReader.Read())
            {
                var buyer = buyerReader.GetString(0);
                var email = buyerReader.GetString(1);
                var name = buyerReader.GetString(2);

                if (buyer.Trim().Length > 0)
                {
                    employees.Add(new Employee(buyer, name, email));
                }
            }

            buyerReader.Close();

            return employees;
        }

        public List<PurchaseOrder> GetPurchaseOrdersToNotify()
        {
            if (dbConnection == null)
            {
                throw new Exception("GetPurchasOrdersToNotify - No Database Connection");
            }

            List<PurchaseOrder> pos = new List<PurchaseOrder>();
            var cmd = new OdbcCommand(POQueryString, dbConnection);
            var reader = cmd.ExecuteReader();

            var buyerColumn = reader.GetOrdinal("buyer");
            var poNumColumn = reader.GetOrdinal("ponum");
            var jobNumberColumn = reader.GetOrdinal("defaultjobnum");
            var binColumn = reader.GetOrdinal("user_1");
            var receivedByColumn = reader.GetOrdinal("user_2");
            var receivedOnDateColumn = reader.GetOrdinal("user_3");
            var vendorNameColumn = reader.GetOrdinal("vendorName");
            var notesColumn = reader.GetOrdinal("user_6");
            var workOrderNumberColumn = reader.GetOrdinal("defaultworkorder");
            var attachmentIdColumn = reader.GetOrdinal("attachid");

            while (reader.Read())
            {
                var purchaseOrderNumber = reader.GetString(poNumColumn);
                var receivedBy = reader.GetString(receivedByColumn);
                var bin = reader.GetString(binColumn);
                var receivedOnDate = reader.GetDate(receivedOnDateColumn).ToShortDateString();
                var buyer = reader.GetString(buyerColumn).ToLower();
                var vendor = reader.GetString(vendorNameColumn);
                var notes = reader.GetString(notesColumn);
                var workOrderNumber = reader.GetString(workOrderNumberColumn);
                var jobNumber = reader.GetString(jobNumberColumn);
                var attachmentId = reader.GetString(attachmentIdColumn);
                //                    var job = queries.GetEmailBodyInformation(reader.GetString(jobNumberColumn), purchaseOrderNumber, workOrderNumber);
                //                    var buyerEmployee = GetEmployeeInformation(EmployeeEmailAddresses, buyer);
                //                    var projectManagerEmployee = job.ProjectManagerName.Length > 0 ? GetEmployeeInformation(EmployeeEmailAddresses, job.ProjectManagerName) : new Employee();
                var attachments = GetAttachmentsForPO(attachmentId);

                pos.Add(new PurchaseOrder(purchaseOrderNumber, receivedBy, bin, receivedOnDate, buyer, vendor, notes, workOrderNumber, string.Empty, string.Empty, jobNumber, attachments));
            }

            reader.Close();

            return pos;
        }
        
        public bool MarkPOAsNotified(string poNum)
        {
            return MarkPOAsNotified(dbConnection, poNum);
        }

        private bool MarkPOAsNotified(OdbcConnection cn, string poNum)
        {
            bool result = true;

            try
            {
                var updateCommand = string.Format("update icpo set \"user_5\" = 1 where icpo.ponum = '{0}'", poNum);
                var cmd = new OdbcCommand(updateCommand, cn);
                cmd.ExecuteNonQuery();
            }
            catch
            {
                log.Error(string.Format("Error marking PO num {0} as notified.", poNum));
                result = false;
            }

            return result;
        }

        public bool MarkQuoteAsNotified(Quote quote)
        {
            bool result = true;

            try
            {
                log.Info("MarkQuoteAsNotified");
                var updateCommand = string.Format("update dpwoassign set \"user_18\" = 1 where workorder = '{0}' and ticketnum = '{1}' and person = '{2}' and user_15 = 1 and user_18 = 0", quote.WorkOrder, quote.WorkTicket, quote.ServiceTech);
                log.Info(updateCommand);
                var cmd = new OdbcCommand(updateCommand, dbConnection);
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                log.Error(string.Format("Error marking Quote for work order {0} as notified.", quote.WorkOrder));
                log.Error(ex.Message);
                log.Error(ex.InnerException.Message);
                result = false;
            }

            return result;
        }

        public List<Quote> GetQuotesNeeded()
        {
            if (dbConnection == null)
            {
                throw new Exception("GetQuotesNeeded - No Database Connection");
            }

            var quotesNeeded = new List<Quote>();

            var cmd = new OdbcCommand(QuotesNeededQueryString, dbConnection);
            var reader = cmd.ExecuteReader();

            var workOrderColumn = reader.GetOrdinal("workorder");
            var workTicketColumn = reader.GetOrdinal("ticketnum");
            var customerNameColumn = reader.GetOrdinal("customername");
            var siteNumColumn = reader.GetOrdinal("sitenum");
            var siteNameColumn = reader.GetOrdinal("sitename");
            var descriptionOfWorkColumn = reader.GetOrdinal("workorderdescription");
            var ticketNoteColumn = reader.GetOrdinal("ticketnote");
            var makeColumn = reader.GetOrdinal("manufacturer");
            var modelColumn = reader.GetOrdinal("model");
            var serialNumberColumn = reader.GetOrdinal("serialnumber");
            var serviceTechColumn = reader.GetOrdinal("servicetech");
            var workDateColumn = reader.GetOrdinal("workdate");

            while (reader.Read())
            {
                var workOrder = reader.GetString(workOrderColumn);
                var workTicket = reader.GetString(workTicketColumn);
                var customerName = reader.GetString(customerNameColumn);
                var siteName = reader.GetString(siteNameColumn);
                var descriptionOfWork = reader.GetString(descriptionOfWorkColumn);
                var ticketNote = reader.GetString(ticketNoteColumn);
                var make = reader.GetString(makeColumn);
                var model = reader.GetString(modelColumn);
                var serialNumber = reader.GetString(serialNumberColumn);
                var serviceTech = reader.GetString(serviceTechColumn);
                var workDate = reader.GetDate(workDateColumn).ToShortDateString();
                log.Info("workDate: " + workDate);

                quotesNeeded.Add(new Quote(workOrder, workTicket, customerName, siteName, descriptionOfWork, ticketNote, serviceTech, make, model, serialNumber, workDate));
            }

            return quotesNeeded;
        }

        public List<AgedReceivable> GetAgedReceivables()
        {
            var testList = new List<AgedReceivable>();
            var ar = new AgedReceivable("Invoice Number", "Customer Name", DateTime.Now, 100.00M);
            testList.Add(ar);
            testList.Add(ar);
            testList.Add(ar);
            testList.Add(ar);
            testList.Add(ar);

            return testList;

//            return new List<AgedReceivable>();
        }
    }
}
