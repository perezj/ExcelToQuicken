using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Transactions;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace BanortePDFToExcel
{
    public class PdfExtractor
    {

        bool started = false;

        const double DATE_MAX_X = 100;
        const double DESC_MAX_X = 400;
        const double AMOUNT_MAX_X = 520;

        private bool ContainsHeader(string text)
        {
            var normalized = Util.Normalize(text);

            return normalized.Contains("FECHA") &&
                   normalized.Contains("DESCRIPCIÓN") &&
                   normalized.Contains("MONTO") &&
                   normalized.Contains("SALDO");
        }

        private List<List<UglyToad.PdfPig.Content.Word>> GroupWordsIntoRows(List<UglyToad.PdfPig.Content.Word> words)
        {
            return words
                .GroupBy(w => Math.Round(w.BoundingBox.Bottom, 1))
                .OrderByDescending(g => g.Key)
                .Select(g => g.OrderBy(w => w.BoundingBox.Left).ToList())
                .ToList();
        }
        private bool RowHasAmount(List<Word> row)
        {
            var text = string.Join(" ", row.Select(w => w.Text));

            return Regex.IsMatch(text, @"\d{1,3}(?:,\d{3})*\.\d{2}");
        }

        private string BuildRowText(List<UglyToad.PdfPig.Content.Word> row)
        {
            return string.Join(" ", row.Select(w => w.Text));
        }

        private (string date, string desc, string amount, string balance) ExtractRowColumns(
            List<Word> row,
            double dateMax,
            double descMax,
            double amountMax)
        {
            string date = "";
            string desc = "";
            string amount = "";
            string balance = "";

            foreach (var w in row)
            {
                var text = w.Text;
                var x = w.BoundingBox.Left;

                // Detect date using pattern
                if (string.IsNullOrEmpty(date))
                {
                    var match = Regex.Match(text, @"^(\d{2}-[A-Z]{3}-\d{2})(.*)");

                    if (match.Success)
                    {
                        date = match.Groups[1].Value;

                        var remainder = match.Groups[2].Value;
                        if (!string.IsNullOrWhiteSpace(remainder))
                            desc += remainder + " ";

                        continue;
                    }
                }

                if (x < descMax)
                    desc += text + " ";
                else if (x < amountMax)
                    amount += text + " ";
                else
                    balance += text + " ";
            }

            return (date.Trim(), desc.Trim(), amount.Trim(), balance.Trim());
        }


        private Transaction? ParseTransactionFromColumns(string dateText, string desc, string amountText, string balanceText)
        {
            if (string.IsNullOrWhiteSpace(balanceText))
                return null;

            // Parse date
            var match = Regex.Match(dateText, @"\d{2}-[A-Z]{3}-\d{2}");
            if (!match.Success)
                return null;

            var parts = match.Value.Split('-');
            int day = int.Parse(parts[0]);
            int month = SpanishMonths[parts[1]];
            int year = 2000 + int.Parse(parts[2]);

            var date = new DateTime(year, month, day).ToString("MM-dd-yy");

            // Parse balance
            var balance = decimal.Parse(balanceText.Replace(",", ""));

            decimal? amount = null;

            if (!string.IsNullOrWhiteSpace(amountText))
            {
                var amtMatch = Regex.Match(amountText, @"\d{1,3}(?:,\d{3})*\.\d{2}");
                if (amtMatch.Success)
                    amount = decimal.Parse(amtMatch.Value.Replace(",", ""));
            }

            return new Transaction
            {
                Date = date,
                Description = desc,
                Amount = amount,
                Balance = balance
            };
        }

        double DetectDateMax(List<List<Word>> rows)
        {
            foreach (var row in rows)
            {
                foreach (var w in row)
                {
                    if (Regex.IsMatch(w.Text, @"^\d{2}-[A-Z]{3}-\d{2}"))
                    {
                        // return right boundary of the date column
                        return w.BoundingBox.Right + 2; // small padding
                    }
                }
            }

            return 100; // fallback
        }
        public List<Account> Extract(string pdfPath)
        {
            bool foundStart = false;
            string buffer = "";
            Account? currentAccount = null;
            var accountsMap = new Dictionary<string, Account>();


            using (var document = PdfDocument.Open(pdfPath))
            {

                for (int p = 1; p <= document.NumberOfPages; p++)
                {
                    var page = document.GetPage(p);
                    var words = page.GetWords().ToList();
                    var rows = GroupWordsIntoRows(words);

                    var xPositions = CollectCandidateXPositions(rows);
                    var clusters = ClusterXPositions(xPositions);

                    var columnResult = DetectColumns(clusters);
                    if (columnResult == null)
                    {
                        continue;
                    }
                    var (dateX, descX, amountX, balanceX) = columnResult.Value;

                    // This is using headers
                    var (dateMax, descMax, amountMax) = BuildColumnBoundaries(dateX, descX, amountX, balanceX);
                    
                    //dateMax = DetectDateMax(rows);

                    Console.WriteLine("---- PAGE TEXT ----");
                    Console.WriteLine(words);
                    Console.WriteLine("----- END --------------");

                    // Normalize spacing
                    //    text = Regex.Replace(text, @"\s+", " ");

                    string currentAccountName = "";
                    var transactionBlocks = new List<TransactionBlock>();
                    TransactionBlock currentBlock = null;

                    for (int i = 0; i < rows.Count; i++)
                    {
                        var row = rows[i];
                        var rowText = BuildRowText(rows[i]);


                        if (!foundStart)
                        {
                            if (ContainsHeader(rowText))
                            {
                                foundStart = true;

                                if (i > 0)
                                {
                                    currentAccountName = BuildRowText(rows[i - 1]).Trim();
                                }
                            }
                            else
                            {
                                continue;
                            }
                        }


                        var (date, desc, amountText, balanceText) =
                            ExtractRowColumns(row, dateMax, descMax, amountMax);

                        if (string.IsNullOrWhiteSpace(date) && string.IsNullOrWhiteSpace(desc))
                            continue;

                        if (rowText.Contains("SIN MOVIMIENTOS", StringComparison.OrdinalIgnoreCase))
                        {
                            if (currentBlock != null)
                            {
                                transactionBlocks.Add(currentBlock);
                                currentBlock = null;
                            }
                            continue;
                        }


                        bool isSaldo = IsSaldoAnterior(desc);

                        // 🔥 NORMAL transaction detection (2 amounts at end)
                        bool hasTwoAmounts = Regex.IsMatch(rowText,
                            @"\d{1,3}(?:,\d{3})*\.\d{2}\s+\d{1,3}(?:,\d{3})*\.\d{2}$");

                        bool isTransactionStart = isSaldo || hasTwoAmounts;

                        if (isTransactionStart)
                        {
                            if (currentBlock != null)
                                transactionBlocks.Add(currentBlock);

                            currentBlock = new TransactionBlock
                            {
                                Rows = new List<List<Word>> { row },
                                AccountName = currentAccountName
                            };
                        }
                        else
                        {
                            if (ContainsHeader(rowText))
                            {
                                if (i > 0)
                                {
                                    currentAccountName = BuildRowText(rows[i - 1]).Trim();
                                }
                                
                                // We are at a header with this rowText so finish adding the block.
                                if (currentBlock != null)
                                {
                                    
                                    if (currentBlock.Rows.Count >= 2)
                                    {
                                        var firstRow = currentBlock.Rows.First();
                                        // Most likely it has the account name 
                                        // Check if the first word is not in the X position of description then its centered and we dont want it.
                                        if (currentBlock.Rows.Last()[0].BoundingBox.Left > firstRow[0].BoundingBox.Left)
                                        {
                                            // Remove the last row as it is not needed
                                            currentBlock.Rows.RemoveAt(currentBlock.Rows.Count - 1);
                                        }
                                    }
                                    transactionBlocks.Add(currentBlock);
                                    currentBlock = null;
                                }
                                    
                            }
                            else
                            {
                                // 🔥 continuation row (multi-line description)
                                if (currentBlock != null)
                                    currentBlock.Rows.Add(row);
                            }
                            
                        }



                        //// Try parse date
                        //var dateMatch = Regex.Match(date, @"\d{2}-[A-Z]{3}-\d{2}");
                        //if (!dateMatch.Success)
                        //    continue;

                        //var parsed = ParseTransactionFromColumns(date, desc, amountText, balanceText);

                        //if (parsed != null)
                        //    currentAccount.Transactions.Add(parsed);
                    }
                    // add last block
                    if (currentBlock != null)
                        transactionBlocks.Add(currentBlock);

                    

                    foreach (var block in transactionBlocks)
                    {
                        if (string.IsNullOrWhiteSpace(block.AccountName))
                            continue;

                        if (!accountsMap.ContainsKey(block.AccountName))
                        {
                            accountsMap[block.AccountName] = new Account
                            {
                                Name = block.AccountName
                            };
                        }

                        var account = accountsMap[block.AccountName];
                        var firstRow = block.Rows[0];

                        var (date, desc, amountText, balanceText) = ExtractRowColumns(firstRow, dateMax, descMax, amountMax);

                        // Merge description
                        var fullText = string.Join(" ",
                            block.Rows.SelectMany(r => r.Select(w => w.Text)));

                        fullText = Regex.Replace(fullText, @"^\d{2}-[A-Z]{3}-\d{2}", "").Trim();
                        fullText = Regex.Replace(fullText, @"\d{1,3}(?:,\d{3})*\.\d{2}", "").Trim();

                        var dateMatch = Regex.Match(date, @"\d{2}-[A-Z]{3}-\d{2}");
                        if (!dateMatch.Success)
                            continue;

                        var parsed = ParseTransactionFromColumns(date, fullText, amountText, balanceText);

                        if (parsed != null)
                            account.Transactions.Add(parsed);
                    }

                    
                    foundStart = false;// Reset on each page
                    // Split by date pattern like 01-SEP-25
                    //var parts = Regex.Split(text, @"(?=\d{2}-[A-Z]{3}-\d{2})");
                    //01/09/2025 or 01-SEP-25
                    //var parts = Regex.Split(text, @"(?=\d{2}[-/][A-Z0-9]{3,5}[-/]\d{2,4})");

                    //foreach (var part in parts)
                    //{

                    //    if (part.Contains("Inversión Enlace Negocios PFAE"))
                    //    {
                    //        // Exit we have reach the end
                    //        break;
                    //    }
                    //    var tx = ParseTransaction(part);

                    //    if (tx != null)
                    //        transactions.Add(tx);
                    //}


                }


            }
            var  accounts = accountsMap.Values.ToList();
            ;
            Console.WriteLine("Finish Extraction..");
            // validate transaction for each account and fix amounts if they don't match balance changes
            foreach (var account in accounts)
            {
                ValidateAndFixAmounts(account.Transactions);
            }


            return accounts;
        }


        private bool IsSaldoAnterior(string description)
        {
            return description.Contains("SALDO ANTERIOR", StringComparison.OrdinalIgnoreCase)
                   || description.Contains("SALDO INICIAL", StringComparison.OrdinalIgnoreCase);
        }
        private List<double> ClusterXPositions(List<double> positions, double tolerance = 15)
        {
            var clusters = new List<List<double>>();

            foreach (var x in positions.OrderBy(x => x))
            {
                var cluster = clusters.FirstOrDefault(c => Math.Abs(c.Average() - x) < tolerance);

                if (cluster == null)
                {
                    clusters.Add(new List<double> { x });
                }
                else
                {
                    cluster.Add(x);
                }
            }

            // Return cluster centers
            return clusters.Select(c => c.Average()).OrderBy(x => x).ToList();
        }
        private (double dateMax, double descMax, double amountMax) BuildColumnBoundaries(
            double dateX, double descX, double amountX, double balanceX)
        {
            return (
                (dateX + descX) / 2,
                (descX + amountX) / 2,
                (amountX + balanceX) / 2
            );
        }
        private (double dateX, double descX, double amountX, double balanceX)? DetectColumns(List<double> clusterCenters)
        {
            if (clusterCenters.Count < 4)
                return null; // 🔥 no exception


            // Left → right
            var ordered = clusterCenters.OrderBy(x => x).ToList();

            double dateX = ordered[0];
            double descX = ordered[1];
            double amountX = ordered[^2];
            double balanceX = ordered[^1];

            return (dateX, descX, amountX, balanceX);
        }

        private List<double> CollectCandidateXPositions(List<List<Word>> rows)
        {
            var xPositions = new List<double>();

            foreach (var row in rows)
            {
                var rowText = string.Join(" ", row.Select(w => w.Text));

                // Only rows with date → likely transaction rows
                if (!Regex.IsMatch(rowText, @"\d{2}-[A-Z]{3}-\d{2}"))
                    continue;

                foreach (var word in row)
                {
                    xPositions.Add(word.BoundingBox.Left);
                }
            }

            return xPositions;
        }

        private string CleanBanorteNoise(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return text;

            // 🔥 Remove IVA blocks like: IVA: 0000000000000.00
            text = Regex.Replace(text, @"IVA:\s*0+\.00", "", RegexOptions.IgnoreCase);

            // Optional: clean extra spaces after removal
            text = Regex.Replace(text, @"\s{2,}", " ");

            return text.Trim();
        }
        private bool IsNewDateAtEnd(string text)
        {
            // Detect date at the END of buffer
            return Regex.IsMatch(text,
                @"(\d{2}-[A-Z]{3}-\d{2}|\d{2}/\d{2}/\d{2,4})$",
                RegexOptions.IgnoreCase);
        }
        private bool HasDateNearStart(string text)
        {
            var slice = text.Substring(0, Math.Min(25, text.Length));

            return Regex.IsMatch(slice, @"\d{2}-[A-Z]{3}-\d{2}");
        }

        private bool TryParseIfComplete(string text, out Transaction tx)
        {
            tx = null;

            if (string.IsNullOrWhiteSpace(text))
                return false;

            // Must have at least 1 amount
            var matches = Regex.Matches(text, @"\d{1,3}(?:,\d{3})*\.\d{2}");

            if (matches.Count == 0)
                return false;

            // Must have date near start
            if (!HasDateNearStart(text))
                return false;

            tx = ParseTransaction(text);

            return tx != null;
        }

        private string ExtractLastDate(string text)
        {
            var match = Regex.Match(text,
                @"(\d{2}-[A-Z]{3}-\d{2}|\d{2}/\d{2}/\d{2,4})$",
                RegexOptions.IgnoreCase);

            return match.Success ? match.Value : "";
        }
        private string NormalizeDates(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return text;

            var matches = Regex.Matches(text, @"\d{2}-[A-Z]{3}-\d{2}", RegexOptions.IgnoreCase);

            if (matches.Count <= 1)
                return text;

            // Keep the FIRST date, clean the rest
            for (int i = 1; i < matches.Count; i++)
            {
                var m = matches[i];

                // Replace with space (important!)
                text = text.Remove(m.Index, m.Length)
                    .Insert(m.Index, " ");
            }

            return text;
        }
        private string FixAmountBoundaries(string text)
        {
            // Fix: 28.80198.59 → 28.80 198.59
            text = Regex.Replace(text,
                @"(\d{1,3}(?:,\d{3})*\.\d{2})(\d{1,3}(?:,\d{3})*\.\d{2})",
                "$1 $2");

            return text;
        }
        private void ExtractTransactionsFromPage(string text, List<Transaction> transactions, ref string buffer)
        {
            text = Regex.Replace(text, @"\s+", " ");

            for (int i = 0; i < text.Length; i++)
            {
                buffer += text[i];

                // Clean noise progressively
                var current = CleanBanorteNoise(buffer);

                if (IsNewDateAtEnd(current))
                {
                    // 🔥 NEW GUARD: must have at least 1 amount before splitting
                    var amountMatches = Regex.Matches(current, @"\d{1,3}(?:,\d{3})*\.\d{2}");

                    if (amountMatches.Count == 0)
                    {
                        // ❌ Not a real transaction yet → keep buffering
                        continue;
                    }

                    string nextDate = ExtractLastDate(current);
                    // Remove the end date of the line to jsut have a clean row starting with date ending with amounts.
                    string cleanedCurrent = current.Substring(0, current.Length - nextDate.Length).Trim();

                    cleanedCurrent = CleanBanorteNoise(cleanedCurrent);
                    // However we can have "date"s inside the description
                    cleanedCurrent = NormalizeDates(cleanedCurrent);
                    // Add spaces.
                    cleanedCurrent = FixAmountBoundaries(cleanedCurrent);

                    if (TryParseIfComplete(cleanedCurrent, out var tx))
                    {
                        transactions.Add(tx);
                    }


                    // ✅ Start new buffer with this date
                    buffer = nextDate;
                    continue;
                }

                // 🔥 STOP if new account starts (second SALDO ANTERIOR)
                if (current.Contains("SALDO ANTERIOR", StringComparison.OrdinalIgnoreCase))
                {
                    if (transactions.Count > 0)
                    {
                        buffer = "";
                        break;
                    }
                }

                bool isEndTransaction = EndOfTransactionPattern.IsMatch(current);
                bool isSaldoAnterior = EndOfSaldoAnteriorPattern.IsMatch(current);

                if (isEndTransaction || isSaldoAnterior)
                {

                    // Extra safety: must contain a date near start
                    if (HasDateNearStart(current))
                    {
                        var tx = ParseTransaction(current.Trim());

                        if (tx != null)
                        {
                            transactions.Add(tx);
                        }
                    }

                    buffer = ""; // 🔥 reset for next transaction
                }

                // Safety: avoid runaway buffer
                if (buffer.Length > 2000)
                {
                    buffer = "";
                }
            }
        }
        private void ValidateAndFixAmounts(List<Transaction> transactions)
        {
            Console.WriteLine("Verifying amounts...");
            for (int i = 1; i < transactions.Count; i++)
            {
                var prev = transactions[i - 1];
                var current = transactions[i];

                decimal diff = current.Balance - prev.Balance;

                decimal amount = Math.Abs(diff);

                current.Amount = amount;

                if (diff > 0)
                {
                    current.Credit = amount;
                    current.Debit = null;
                }
                else
                {
                    current.Debit = amount;
                    current.Credit = null;
                }
            }
            Console.WriteLine("Verifying amounts complete!");
        }



        private static readonly Regex EndOfSaldoAnteriorPattern =
            new Regex(@"SALDO ANTERIOR.*\d{1,3}(?:,\d{3})*\.\d{2}$", RegexOptions.IgnoreCase);


        private static readonly Regex EndOfTransactionPattern =
            new Regex(@"\d{1,3}(?:,\d{3})*\.\d{2}\s+\d{1,3}(?:,\d{3})*\.\d{2}$");

        // Heuristic to check if text ends with two amounts (amount and balance)
        //private bool LooksLikeTransaction(string text)
        //{
        //    if (string.IsNullOrWhiteSpace(text))
        //        return false;

        //    text = text.Trim();

        //    // Must start with date
        //    if (!Regex.IsMatch(text, @"^\d{2}-[A-Z]{3}-\d{2}"))
        //        return false;


        //    var matches = Regex.Matches(text, @"\d{1,3}(,\d{3})*\.\d{2}");

        //    if (matches.Count == 0)
        //        return false;

        //    var last = matches[matches.Count - 1];

        //    // Ensure last amount is near the end
        //    int endDistance = text.Length - last.Index;
        //    if (endDistance > 20)
        //        return false;

        //    // 🔥 CASE 1: Only 1 amount → SALDO ANTERIOR
        //    if (matches.Count == 1)
        //    {
        //        if (text.Contains("SALDO ANTERIOR", StringComparison.OrdinalIgnoreCase))
        //            return true;

        //        return false;
        //    }

        //    // 🔥 CASE 2: Normal transaction (2+ amounts)
        //    var secondLast = matches[matches.Count - 2];

        //    int gap = last.Index - (secondLast.Index + secondLast.Length);
        //    if (gap > 5)
        //        return false;

        //    return true;
        //}


        static readonly Dictionary<string, int> SpanishMonths = new()
        {
            { "ENE", 1 },
            { "FEB", 2 },
            { "MAR", 3 },
            { "ABR", 4 },
            { "MAY", 5 },
            { "JUN", 6 },
            { "JUL", 7 },
            { "AGO", 8 },
            { "SEP", 9 },
            { "OCT", 10 },
            { "NOV", 11 },
            { "DIC", 12 }
        };

        static string FixBrokenAmounts(string text)
        {
            // Fix cases like: 14:14:38320.00 → 14:14:38 320.00
            text = Regex.Replace(text,
                @"(\d{2}:\d{2}:\d{2})(\d{1,3}(?:,\d{3})*\.\d{2})",
                "$1 $2");

            // Fix cases like: 320.003,392.64 → 320.00 3,392.64
            text = Regex.Replace(text,
                @"(\d{1,3}(?:,\d{3})*\.\d{2})(\d{1,3},\d{3}\.\d{2})",
                "$1 $2");

            return text;
        }


        static Transaction? ParseTransaction(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return null;

            text = FixBrokenAmounts(text);

            text = text.Trim();

            var dateMatch = Regex.Match(text, @"\d{2}-[A-Z]{3}-\d{2}");
            if (!dateMatch.Success)
                return null;

            string rawDate = dateMatch.Value.ToUpper();

            var parts = rawDate.Split('-');

            int day = int.Parse(parts[0]);
            int month = SpanishMonths[parts[1]];
            int year = 2000 + int.Parse(parts[2]);

            DateTime parsedDate = new DateTime(year, month, day);
            string date = parsedDate.ToString("MM-dd-yy");

            text = text.Substring(dateMatch.Index + dateMatch.Length).Trim();
            text = Regex.Replace(text, @"\d{2}:\d{2}:\d{2}", "");
            text = text.Replace("$", "");

            var amountMatches = Regex.Matches(text, @"\d{1,3}(,\d{3})*\.\d{2}");

            var validMatches = new List<Match>();
            var validAmounts = new List<decimal>();

            foreach (Match m in amountMatches)
            {
                string raw = m.Value.Replace(",", "");

                if (raw.Length > 10)
                    continue;

                if (Regex.IsMatch(raw, @"^0{4,}\.\d{2}$"))
                    continue;

                if (raw.StartsWith("0000"))
                    continue;

                if (decimal.TryParse(raw, out var value))
                {
                    validMatches.Add(m);
                    validAmounts.Add(value);
                }
            }

            if (validAmounts.Count == 0)
                return null;

            decimal balance;
            decimal? amount = null;

            if (validAmounts.Count == 1)
            {
                // 🔥 SALDO ANTERIOR
                balance = validAmounts[0];

                // Remove only that match
                text = text.Replace(validMatches[0].Value, "");
            }
            else
            {
                // Always take LAST TWO
                var lastMatch = validMatches[^1];
                var secondLastMatch = validMatches[^2];

                balance = validAmounts[^1];
                amount = validAmounts[^2];

                // Remove ONLY those two
                text = text.Replace(lastMatch.Value, "");
                text = text.Replace(secondLastMatch.Value, "");
            }

            string description = text.Trim();

            //bool isDeposit =
            //    description.Contains("SPEI RECIBIDO", StringComparison.OrdinalIgnoreCase) ||
            //    description.Contains("DEPOSITO", StringComparison.OrdinalIgnoreCase) ||
            //    description.Contains("TRANSFERENCIA", StringComparison.OrdinalIgnoreCase) ||
            //    description.Contains("ABONO", StringComparison.OrdinalIgnoreCase);

            return new Transaction
            {
                Date = date,
                Description = description,
                Balance = balance,
                Amount = amount,
                Credit = null,// isDeposit ? amount : null,
                Debit = null,//!isDeposit ? amount : null,
            };
        }
    }
}