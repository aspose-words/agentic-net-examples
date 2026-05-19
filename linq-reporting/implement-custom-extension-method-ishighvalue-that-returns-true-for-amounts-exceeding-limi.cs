using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExtensionDemo
{
    // Extension method used in the LINQ Reporting template.
    public static class Extensions
    {
        // Returns true if the amount exceeds the specified limit.
        public static bool IsHighValue(this decimal amount, decimal limit) => amount > limit;
    }

    // Data model for the report.
    public class Item
    {
        public decimal Amount { get; set; } = 0m;
    }

    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Create a blank document that will serve as the template.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build the LINQ Reporting template.
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("Amount: <<[item.Amount]>>");
            // Call the custom extension method IsHighValue with a limit of 1000.
            builder.Writeln("<<if [item.Amount.IsHighValue(1000)]>>High Value<</if>>");
            builder.Writeln("<</foreach>>");

            // Prepare sample data.
            ReportModel model = new ReportModel();
            model.Items.Add(new Item { Amount = 500m });
            model.Items.Add(new Item { Amount = 1500m });
            model.Items.Add(new Item { Amount = 2500m });

            // Configure the reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.AllowMissingMembers; // Enable extension method usage.
            engine.KnownTypes.Add(typeof(Extensions)); // Register the type containing the extension method.

            // Build the report using the model as the root data source named "model".
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            doc.Save("Report.docx");
        }
    }
}
