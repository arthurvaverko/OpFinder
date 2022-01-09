using System.Collections.Generic;
using System.Security.Cryptography;

namespace OpFinder
{
    public class ResultRow
    {
        public string ClientId { get; set; }
        public string TempOpId{ get; set; }
        public string EngagementOrInvoiceId { get; set; }
    }
}