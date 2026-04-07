using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model with a collection of items.
    public class ReportModel
    {
        // Initialize the collection to avoid nullable warnings.
        public List<Item> Items { get; set; } = new();

        // Populate the collection with sample data.
        public ReportModel()
        {
            Items.Add(new Item { Name = "Alpha" });
            Items.Add(new Item { Name = "Beta" });
            Items.Add(new Item { Name = "Gamma" });
        }
    }

    // Individual item class.
    public class Item
    {
        public string Name { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -------------------------------------------------
            // 1. Create the template document programmatically.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a line that uses ElementAt to fetch the last item.
            // The expression calculates the index as Items.Count - 1.
            builder.Writeln("Last item: <<[model.Items.ElementAt(model.Items.Count - 1).Name]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template and build the report.
            // -------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // Create the data source.
            ReportModel model = new ReportModel();

            // Use the ReportingEngine to populate the template.
            ReportingEngine engine = new ReportingEngine();
            // No special options are required for this simple scenario.
            engine.Options = ReportBuildOptions.None;

            // Build the report; the root object name in the template is "model".
            engine.BuildReport(reportDoc, model, "model");

            // -------------------------------------------------
            // 3. Save the generated report.
            // -------------------------------------------------
            reportDoc.Save(reportPath);
        }
    }
}
