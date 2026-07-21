using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Simple data model used by the template.
    public class ReportModel
    {
        // Title of the heading.
        public string Title { get; set; } = "Dynamic Heading Example";
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            string templatePath = Path.Combine(Environment.CurrentDirectory, "template.docx");
            string outputPath = Path.Combine(Environment.CurrentDirectory, "report.docx");

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Set the desired font size for the heading.
            builder.Font.Size = 14;

            // Insert a heading paragraph with a LINQ Reporting tag.
            // The <<[model.Title]>> tag will be replaced with the Title property of the model at runtime.
            builder.Writeln("<<[model.Title]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare the data source.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel
            {
                Title = "Report Heading – Font Size 14pt"
            };

            // -----------------------------------------------------------------
            // 3. Load the template and build the report.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();

            // Build the report using the model as the root data source named "model".
            engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            doc.Save(outputPath);
        }
    }
}
