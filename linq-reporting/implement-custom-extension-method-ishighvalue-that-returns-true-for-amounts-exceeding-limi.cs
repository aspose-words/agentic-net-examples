using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Extension method used in the LINQ Reporting template.
    public static class Extensions
    {
        // Returns true if the amount exceeds the specified limit.
        public static bool IsHighValue(this decimal amount, decimal limit) => amount > limit;
    }

    // Data model exposed to the template.
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public decimal Amount { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create the template document with LINQ Reporting tags.
            // -----------------------------------------------------------------
            var template = new Document();
            var builder = new DocumentBuilder(template);

            // Begin a foreach loop over Items.
            builder.Writeln("<<foreach [item in Items]>>");
            // Write the amount and evaluate the custom extension method.
            builder.Writeln("Amount: <<[item.Amount]>>  High: <<[item.Amount.IsHighValue(1000)]>>");
            // End the foreach loop.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and prepare the data source.
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);

            var model = new ReportModel
            {
                Items = new List<Item>
                {
                    new Item { Amount = 500m },
                    new Item { Amount = 1500m },
                    new Item { Amount = 2500m }
                }
            };

            // -----------------------------------------------------------------
            // 3. Build the report using ReportingEngine.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine
            {
                // Allow the engine to resolve extension methods.
                Options = ReportBuildOptions.AllowMissingMembers
            };

            // The root object name in the template is "model".
            engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            const string reportPath = "Report.docx";
            doc.Save(reportPath);
        }
    }
}
