using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Model class exposed to the template.
    public class ReportModel
    {
        // Initialize the collection to avoid nullable warnings.
        public List<string> Items { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data.
            var model = new ReportModel();
            model.Items.AddRange(new[] { "Apple", "Banana", "Cherry", "Date", "Elderberry" });

            // -----------------------------------------------------------------
            // Step 1: Create the template document programmatically.
            // -----------------------------------------------------------------
            var template = new Document();
            var builder = new DocumentBuilder(template);

            // Insert a LINQ Reporting tag that fetches the third item (index 2) using ElementAt.
            builder.Writeln("Third item: <<[model.Items.ElementAt(2)]>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // Step 2: Load the template and build the report.
            // -----------------------------------------------------------------
            var loadedTemplate = new Document(templatePath);
            var engine = new ReportingEngine();

            // Build the report using the model and the root name "model".
            engine.BuildReport(loadedTemplate, model, "model");

            // Save the generated report.
            const string reportPath = "Report.docx";
            loadedTemplate.Save(reportPath);
        }
    }
}
