using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Data model used by the LINQ Reporting template.
    public class ReportModel
    {
        // Nullable discount value.
        public decimal? Discount { get; set; }

        // Nullable tax value.
        public decimal? Tax { get; set; }

        // Combined value using the lifted addition operator.
        // If either operand is null, the result is null.
        public decimal? Combined => Discount + Tax;
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a simple Word template with a LINQ Reporting tag.
            // -----------------------------------------------------------------
            string templatePath = "Template.docx";

            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Write a line that will display the combined value.
            builder.Writeln("Combined value: <<[model.Combined]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare sample data.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel
            {
                Discount = 12.5m,
                Tax = 3.75m
                // If you want to test null handling, set either property to null.
            };

            // -----------------------------------------------------------------
            // 3. Load the template and build the report.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            ReportingEngine engine = new ReportingEngine();
            // The root object name used in the template tags is "model".
            engine.BuildReport(reportDoc, model, "model");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            string outputPath = "Report.docx";
            reportDoc.Save(outputPath);

            // Inform the user (optional, not required for non‑interactive execution).
            Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
        }
    }
}
