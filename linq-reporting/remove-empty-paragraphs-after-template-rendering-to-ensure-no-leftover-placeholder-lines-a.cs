using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model used by the LINQ Reporting template.
    public class ReportModel
    {
        // When true the conditional block in the template will be rendered.
        public bool ShowDetails { get; set; } = false;

        // Text displayed inside the conditional block.
        public string Detail { get; set; } = "Sample detail text.";
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare folders for the template and the final report.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            string templatePath = Path.Combine(outputDir, "template.docx");
            string resultPath = Path.Combine(outputDir, "result.docx");

            // 1. Create a template document programmatically.
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("=== Report Start ===");

            // Conditional paragraph – will be removed when ShowDetails is false.
            // The whole paragraph becomes empty after the engine processes the tag,
            // and the RemoveEmptyParagraphs option will delete it.
            builder.Writeln("<<if [model.ShowDetails]>>Details: <<[model.Detail]>> <</if>>");

            builder.Writeln("=== Report End ===");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // 2. Load the template for reporting.
            Document reportDoc = new Document(templatePath);

            // 3. Prepare the data source.
            ReportModel model = new ReportModel
            {
                ShowDetails = false, // Change to true to keep the conditional paragraph.
                Detail = "This text would appear if ShowDetails were true."
            };

            // 4. Configure the ReportingEngine to remove empty paragraphs.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // 5. Build the report using the model as the root object named "model".
            engine.BuildReport(reportDoc, model, "model");

            // 6. Save the final document.
            reportDoc.Save(resultPath);

            // Output the resulting text to the console for verification.
            Console.WriteLine("Report generated. Content:");
            Console.WriteLine(reportDoc.GetText().Trim());
        }
    }
}
