using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model used by the LINQ Reporting template.
    public class ReportModel
    {
        // Initialize the collection to avoid nullable warnings.
        public List<int> Numbers { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create the LINQ Reporting template programmatically.
            // -----------------------------------------------------------------
            var template = new Document();
            var builder = new DocumentBuilder(template);

            // Write a simple list of numbers using a foreach tag.
            builder.Writeln("Numbers:");
            builder.Writeln("<<foreach [n in model.Numbers]>>");
            builder.Writeln("<<[n]>>");
            builder.Writeln("<</foreach>>");

            // Use ElementAt to fetch the second‑to‑last item by calculating the length.
            builder.Writeln("Second to last: <<[model.Numbers.ElementAt(model.Numbers.Count - 2)]>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back for report generation.
            // -----------------------------------------------------------------
            var document = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data source.
            // -----------------------------------------------------------------
            var model = new ReportModel
            {
                Numbers = new List<int> { 10, 20, 30, 40, 50 }
            };

            // -----------------------------------------------------------------
            // 4. Build the report using Aspose.Words LINQ Reporting engine.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine();
            engine.BuildReport(document, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            document.Save(outputPath);
        }
    }
}
