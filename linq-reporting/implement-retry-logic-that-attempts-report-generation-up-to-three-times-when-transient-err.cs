using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

namespace AsposeWordsLinqReportingRetry
{
    // Simple data model for the report.
    public class Order
    {
        public string CustomerName { get; set; } = "John Doe";
        public List<Item> Items { get; set; } = new()
        {
            new Item { Index = 1, Name = "Apple" },
            new Item { Index = 2, Name = "Banana" }
        };
    }

    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "ReportTemplate.docx";
            const string outputPath = "ReportResult.docx";

            // 1. Create a template document with LINQ Reporting tags.
            CreateTemplate(templatePath);

            // 2. Load the template.
            Document doc = new Document(templatePath);

            // 3. Prepare the data source.
            Order order = new();

            // 4. Build the report with retry logic (max 3 attempts).
            bool success = false;
            ReportingEngine engine = new();

            for (int attempt = 1; attempt <= 3 && !success; attempt++)
            {
                try
                {
                    // BuildReport overload that allows referencing the root object name.
                    engine.BuildReport(doc, order, "order");
                    success = true;
                }
                catch (Exception ex) when (IsTransient(ex))
                {
                    Console.WriteLine($"Attempt {attempt} failed with a transient error: {ex.Message}");
                    if (attempt == 3)
                        throw; // rethrow after final attempt

                    // Simple back‑off before retrying.
                    Thread.Sleep(500);
                }
            }

            // 5. Save the generated report.
            if (success)
                doc.Save(outputPath);
        }

        // Determines whether an exception is considered transient.
        private static bool IsTransient(Exception ex)
        {
            // For demonstration, treat all exceptions as transient.
            // In real scenarios, inspect the exception type/message.
            return true;
        }

        // Creates a minimal Word template containing LINQ Reporting tags.
        private static void CreateTemplate(string path)
        {
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert a simple paragraph with a tag that references the root object.
            builder.Writeln("Customer: <<[order.CustomerName]>>");

            // Insert a table that iterates over the Items collection.
            builder.Writeln("<<foreach [item in order.Items]>>");
            Table table = builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Writeln("Index");
            builder.InsertCell();
            builder.Writeln("Name");
            builder.EndRow();

            // Data row.
            builder.InsertCell();
            builder.Writeln("<<[item.Index]>>");
            builder.InsertCell();
            builder.Writeln("<<[item.Name]>>");
            builder.EndRow();

            builder.EndTable();
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            template.Save(path);
        }
    }
}
