using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingLoggingExample
{
    // Simple logger that records custom messages.
    public class ListLogger
    {
        public List<string> Messages { get; } = new();

        public void Log(string message)
        {
            Messages.Add(message);
            Console.WriteLine(message);
        }
    }

    // Data model for the report.
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
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
            // Prepare sample data.
            var model = new ReportModel
            {
                Items = new List<Item>
                {
                    new Item { Index = 1, Name = "Alpha" },
                    new Item { Index = 2, Name = "Beta" },
                    new Item { Index = 3, Name = "Gamma" }
                }
            };

            // Create a template document with LINQ Reporting tags.
            const string templatePath = "Template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("Report generated with LINQ Reporting");
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("Item <<[item.Index]>>: <<[item.Name]>>");
            builder.Writeln("<</foreach>>");

            doc.Save(templatePath);

            // Load the template for report generation.
            var template = new Document(templatePath);

            // Set up a simple logger to capture evaluation details.
            var logger = new ListLogger();

            // Manually log the data that will be used in the report.
            logger.Log("Starting report generation. Data items:");
            foreach (var itm in model.Items)
            {
                logger.Log($"Item Index={itm.Index}, Name={itm.Name}");
            }

            // Build the report.
            var engine = new ReportingEngine();
            engine.BuildReport(template, model, "model");

            // Save the generated report.
            const string outputPath = "Report.docx";
            template.Save(outputPath);

            // Write a summary of logged messages.
            Console.WriteLine();
            Console.WriteLine("Logging summary:");
            foreach (var msg in logger.Messages)
            {
                Console.WriteLine(msg);
            }
        }
    }
}
