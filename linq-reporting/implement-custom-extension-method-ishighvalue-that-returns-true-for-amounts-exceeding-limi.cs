using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public static class Extensions
{
    // Returns true if the transaction amount exceeds the specified limit.
    public static bool IsHighValue(this Transaction transaction, decimal limit)
    {
        return transaction.Amount > limit;
    }
}

// Data model for the report.
public class ReportModel
{
    public List<Transaction> Items { get; set; } = new();
}

// Simple transaction class.
public class Transaction
{
    public decimal Amount { get; set; }
}

public class Program
{
    public static void Main()
    {
        // 1. Create the template document programmatically.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // LINQ Reporting tags: iterate over Items and display amount and high‑value flag.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Amount: <<[item.Amount]>>  High: <<[item.IsHighValue(100)]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        doc.Save(templatePath);

        // 2. Load the template for reporting.
        var reportDoc = new Document(templatePath);

        // 3. Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Transaction>
            {
                new Transaction { Amount = 50m },
                new Transaction { Amount = 150m },
                new Transaction { Amount = 75m },
                new Transaction { Amount = 200m }
            }
        };

        // 4. Build the report using the ReportingEngine.
        var engine = new ReportingEngine
        {
            // Allow the engine to call extension methods like IsHighValue.
            Options = ReportBuildOptions.AllowMissingMembers
        };

        engine.BuildReport(reportDoc, model, "model");

        // 5. Save the generated report.
        reportDoc.Save("ReportOutput.docx");
    }
}
