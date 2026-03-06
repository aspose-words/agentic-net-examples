// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the DOCX template that contains mail‑merge fields (e.g. «Name», «Amount»).
        const string templatePath = @"C:\Docs\ReportTemplate.docx";

        // Load the existing DOCX document.
        Document doc = new Document(templatePath);

        // Sample data source – a collection of records.
        var records = new List<ReportRecord>
        {
            new ReportRecord { Name = "Alice", Amount = 123.45m },
            new ReportRecord { Name = "Bob",   Amount = 678.90m },
            new ReportRecord { Name = "Carol", Amount = 250.00m }
        };

        // Use LINQ to filter the records (e.g., only amounts greater than 200).
        var filtered = records.Where(r => r.Amount > 200).ToArray();

        // Prepare field names and values for the mail‑merge operation.
        string[] fieldNames = { "Name", "Amount" };
        object[] fieldValues = filtered
            .SelectMany(r => new object[] { r.Name, r.Amount })
            .ToArray();

        // Populate the document with the filtered data.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Print the resulting document to the default printer.
        doc.Print();
    }

    // Simple POCO representing a row of data for the report.
    public class ReportRecord
    {
        public string Name { get; set; }
        public decimal Amount { get; set; }
    }
}
