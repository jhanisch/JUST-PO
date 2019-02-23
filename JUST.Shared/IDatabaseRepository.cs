﻿using System;
using System.Collections.Generic;
using System.Net.Mail;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using JUST.Shared.Classes;

namespace JUST.Shared.DatabaseRepository
{
    public interface IDatabaseRepository
    {
        string POQuery { get; }

        List<Attachment> GetAttachmentsForPO(string attachid);

        JobInformation GetEmailBodyInformation(string jobNum, string purchaseOrderNumber, string workOrderNumber);

        List<Employee> GetEmployees();

        List<PurchaseOrder> GetPurchaseOrdersToNotify();

        bool MarkPOAsNotified(string poNum);
    }
}