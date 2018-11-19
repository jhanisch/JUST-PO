using JUST_PONotifier.Classes;
using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace JUST_PONotifier.Queries
{
    public class DatabaseQueries
    {
        private OdbcConnection dbConnection;
        private log4net.ILog log;

        public DatabaseQueries()
        { }

        public DatabaseQueries(OdbcConnection cn, log4net.ILog x)
        {
            dbConnection = cn;
            log = x;
        }

        public List<Attachment> GetAttachmentsForPO(string attachid, string rootPath)
        {
            return GetAttachmentsForPO(dbConnection, attachid, rootPath);
        }

        private List<Attachment> GetAttachmentsForPO(OdbcConnection cn, string attachid, string rootPath)
        {  
            if (cn == null)
            {
                throw new Exception("No Database connection exists");
            }

            List<Attachment> poAttachments = new List<Attachment>();

            var attachmentQuery = "Select path, displayname from icpoattachment where ownerkey = '{0}'";
            var attachmentCmd = new OdbcCommand(string.Format(attachmentQuery, attachid), cn);

            OdbcDataReader attachmentReader = attachmentCmd.ExecuteReader();

            while (attachmentReader.Read())
            {
                var path = attachmentReader.GetString(0);
                var displayName = attachmentReader.GetString(1);

                if (path.Trim().Length > 0)
                {
                    var fullPath = rootPath + path;
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

            OdbcDataReader jobReader = jobCmd.ExecuteReader();

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

        public List<PurchaseOrder> GetUnNotifiedPurchaseOrders()
        {
            List<PurchaseOrder> pos = new List<PurchaseOrder>();

            return pos;
        }

    }
}
