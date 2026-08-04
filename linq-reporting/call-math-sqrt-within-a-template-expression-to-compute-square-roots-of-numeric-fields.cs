using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingSqrt
{
    // Simple data model containing a collection of numeric items.
    public class ReportModel
    {
        // Initialize the collection to avoid nullable warnings.
        public List<Item> Items { get; set; } = new();

        // Sample data constructor.
        public ReportModel()
        {
            Items.Add(new Item { Value = 4 });
            Items.Add(new Item { Value = 9 });
            Items.Add(new Item { Value = 16 });
            Items.Add(new Item { Value = 25 });
        }
    }

    public class Item
    {
        public double Value { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Write a heading.
            builder.Writeln("Square root calculation using Math.Sqrt in LINQ Reporting");
            builder.Writeln();

            // Begin a foreach loop over the Items collection.
            builder.Writeln("<<foreach [item in Items]>>");
            // Output the original value and its square root.
            // Math.Sqrt is a static method; we add System.Math to KnownTypes later.
            builder.Writeln("Value: <<[item.Value]>>, Sqrt: <<[Math.Sqrt(item.Value)]>>");
            // End the foreach loop.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // Prepare the data source.
            ReportModel model = new();

            // Configure the reporting engine.
            ReportingEngine engine = new ReportingEngine();
            // Register System.Math so that static methods can be called from the template.
            engine.KnownTypes.Add(typeof(Math));

            // Build the report. No root name is supplied, so tags reference members directly.
            engine.BuildReport(reportDoc, model);

            // Save the generated report.
            reportDoc.Save(reportPath);
        }
    }
}
