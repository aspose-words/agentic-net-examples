using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingRetryExample
{
    // Sample data model for the report.
    public class ReportModel
    {
        public string CustomerName { get; set; } = "Acme Corp";
        public List<Item> Items { get; set; } = new()
        {
            new Item { Index = 1, Name = "Widget" },
            new Item { Index = 2, Name = "Gadget" },
            new Item { Index = 3, Name = "Doohickey" }
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
            // Register code page provider required by Aspose.Words.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare file paths.
            string templatePath = "template.docx";
            string outputPath = "report.docx";

            // -----------------------------------------------------------------
            // 1. Create the LINQ Reporting template programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Simple header with a placeholder for the customer name.
            builder.Writeln("Customer: <<[model.CustomerName]>>");
            builder.Writeln("Items:");

            // Foreach block to list items.
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("- <<[item.Index]>>: <<[item.Name]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and prepare the data model.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);
            ReportModel model = new();

            // -----------------------------------------------------------------
            // 3. Build the report with retry logic (up to 3 attempts).
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            const int maxAttempts = 3;
            bool success = false;

            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                try
                {
                    // BuildReport returns true if the report was generated without errors.
                    success = engine.BuildReport(doc, model, "model");
                    if (success)
                    {
                        Console.WriteLine($"Report generated successfully on attempt {attempt}.");
                        break;
                    }
                    else
                    {
                        Console.WriteLine($"Report generation failed on attempt {attempt} (engine returned false).");
                    }
                }
                catch (Exception ex) when (IsTransient(ex))
                {
                    // Transient error – log and retry.
                    Console.WriteLine($"Transient error on attempt {attempt}: {ex.Message}");
                }
                catch (Exception ex)
                {
                    // Non‑transient error – abort retries.
                    Console.WriteLine($"Non‑transient error on attempt {attempt}: {ex.Message}");
                    break;
                }
            }

            // -----------------------------------------------------------------
            // 4. Save the final report if generation succeeded.
            // -----------------------------------------------------------------
            if (success)
            {
                doc.Save(outputPath);
                Console.WriteLine($"Report saved to '{outputPath}'.");
            }
            else
            {
                Console.WriteLine("Report generation failed after maximum retry attempts.");
            }
        }

        // Simple heuristic to decide whether an exception is transient.
        private static bool IsTransient(Exception ex)
        {
            // For demonstration, treat IO and network‑related exceptions as transient.
            return ex is IOException || ex is System.Net.WebException;
        }
    }
}
