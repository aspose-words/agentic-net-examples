using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

namespace LinqReportingExample
{
    // Data model
    public class Transaction
    {
        public decimal Amount { get; set; }
        public bool IsHigh { get; set; } = false; // initialized to avoid nullable warnings
    }

    // Extension method
    public static class TransactionExtensions
    {
        public static bool IsHighValue(this Transaction transaction, decimal limit)
        {
            return transaction.Amount > limit;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Sample data
            List<Transaction> transactions = new()
            {
                new Transaction { Amount = 500m },
                new Transaction { Amount = 1500m },
                new Transaction { Amount = 2500m },
                new Transaction { Amount = 750m }
            };

            // Apply extension method to set IsHigh flag
            const decimal limit = 1000m;
            foreach (var t in transactions)
            {
                t.IsHigh = t.IsHighValue(limit);
            }

            // Create template document
            string templatePath = "Template.docx";
            Document templateDoc = new();
            DocumentBuilder builder = new(templateDoc);

            // Title
            builder.Writeln("Transaction Report");
            builder.Writeln();

            // Start foreach block
            builder.Writeln("<<foreach [item in transactions]>>");

            // Table header (only once, before first row)
            builder.Writeln("<<if [item == transactions[0]]>>");
            Table headerTable = builder.StartTable();
            builder.InsertCell();
            builder.Writeln("Amount");
            builder.InsertCell();
            builder.Writeln("High Value");
            builder.EndRow();
            builder.EndTable();
            builder.Writeln("<</if>>");

            // Table row for each item
            Table rowTable = builder.StartTable();
            builder.InsertCell();
            builder.Writeln("<<[item.Amount]>>");
            builder.InsertCell();
            builder.Writeln("<<if [item.IsHigh]>>Yes<</if>><<if [!item.IsHigh]>>No<</if>>");
            builder.EndRow();
            builder.EndTable();

            // End foreach block
            builder.Writeln("<</foreach>>");

            // Save template
            templateDoc.Save(templatePath);

            // Load template for reporting
            Document doc = new(templatePath);
            ReportingEngine engine = new();

            // Build report using the list as the root object named "transactions"
            engine.BuildReport(doc, transactions, "transactions");

            // Save final report
            string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
