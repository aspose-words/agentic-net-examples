using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model for the report.
    public class ReportModel
    {
        public string Title { get; set; } = "Dynamic Heading Example";
    }

    public class Program
    {
        public static void Main()
        {
            // Ensure the output directory exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Paths for the template and the generated report.
            string templatePath = Path.Combine(outputDir, "Template.docx");
            string reportPath = Path.Combine(outputDir, "Report.docx");

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Set the paragraph style to Heading1 and the font size to 14 points.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Font.Size = 14;

            // Insert the LINQ Reporting tag that will be replaced with the model's Title.
            builder.Writeln("<<[model.Title]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare the data source.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel();

            // -----------------------------------------------------------------
            // 3. Build the report using the ReportingEngine.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();

            // BuildReport overload that takes the data source name ("model").
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report.
            reportDoc.Save(reportPath);
        }
    }
}
