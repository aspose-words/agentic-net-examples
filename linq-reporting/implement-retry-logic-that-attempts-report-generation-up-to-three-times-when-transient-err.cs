using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingRetryExample
{
    // Data model classes
    public class Order
    {
        public string CustomerName { get; set; } = string.Empty;
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = string.Empty;
        public double Price { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Create template document programmatically
            var templatePath = "Template.docx";
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln("Items:");
            builder.Writeln("<<foreach [item in order.Items]>>");
            builder.Writeln("- <<[item.Name]>>: $<<[item.Price]>>");
            builder.Writeln("<</foreach>>");
            templateDoc.Save(templatePath);

            // Load the template for report generation
            var doc = new Document(templatePath);

            // Prepare sample data
            var order = new Order
            {
                CustomerName = "John Doe",
                Items = new()
                {
                    new Item { Name = "Apple", Price = 1.20 },
                    new Item { Name = "Banana", Price = 0.80 },
                    new Item { Name = "Cherry", Price = 2.50 }
                }
            };

            // Reporting engine instance
            var engine = new ReportingEngine();

            // Retry logic: up to three attempts for transient errors
            const int maxAttempts = 3;
            bool success = false;
            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                try
                {
                    // Build the report
                    engine.BuildReport(doc, order, "order");
                    success = true;
                    break; // Exit loop on success
                }
                catch (Exception ex) when (IsTransient(ex) && attempt < maxAttempts)
                {
                    // Log transient error and retry
                    Console.WriteLine($"Transient error on attempt {attempt}: {ex.Message}. Retrying...");
                }
                catch (Exception ex)
                {
                    // Non‑transient error or final attempt failure
                    Console.WriteLine($"Report generation failed: {ex.Message}");
                    break;
                }
            }

            if (success)
            {
                // Save the generated report
                var outputPath = "Report.docx";
                doc.Save(outputPath);
                Console.WriteLine($"Report generated successfully: {outputPath}");
            }
        }

        // Simple heuristic to decide if an exception is transient
        private static bool IsTransient(Exception ex)
        {
            // For demonstration, treat all IO and network related exceptions as transient
            return ex is IOException;
        }
    }
}
