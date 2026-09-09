using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model containing a collection of strings.
    public class ReportModel
    {
        // Initialize the list to avoid nullable warnings.
        public List<string> Items { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data with at least three items.
            var model = new ReportModel
            {
                Items = new List<string> { "First", "Second", "Third", "Fourth" }
            };

            // -----------------------------------------------------------------
            // Step 1: Create a Word template programmatically.
            // -----------------------------------------------------------------
            var templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Insert a tag that uses ElementAt to fetch the third item (index 2).
            builder.Writeln("Third item: <<[model.Items.ElementAt(2)]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // Step 2: Load the template and build the report.
            // -----------------------------------------------------------------
            var reportDoc = new Document(templatePath);

            // Use the ReportingEngine to merge the model with the template.
            var engine = new ReportingEngine();
            // The root object name must match the name used in the template tags ("model").
            engine.BuildReport(reportDoc, model, "model");

            // -----------------------------------------------------------------
            // Step 3: Save the generated report.
            // -----------------------------------------------------------------
            var outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
            reportDoc.Save(outputPath);

            // The example finishes execution here. No user interaction is required.
        }
    }
}
