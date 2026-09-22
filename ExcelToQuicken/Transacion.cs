using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using UglyToad.PdfPig.Content;

namespace BanortePDFToExcel
{

    public class Account
    {
        public string Name { get; set; } = "";
        public List<Transaction> Transactions { get; set; } = new();
    }
    public class TransactionBlock
    {
        public List<List<Word>> Rows { get; set; } = new();
        public string AccountName { get; set; } = "";
    }

    public class Transaction
    {
        public string Date { get; set; } = "";
        public string Description { get; set; } = "";
        public decimal? Debit { get; set; }
        public decimal? Credit { get; set; }
        public decimal Balance { get; set; }
        public decimal? Amount { get; set; }
        public string Memo { get; set; } = "";
        public string Category { get; set; } = "";
        public string Tag { get; set; } = "";
    }
}
