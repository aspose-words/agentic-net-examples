using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model used by the template.
    public class ReportModel
    {
        // This property is intentionally empty; its tag will be removed,
        // leaving an empty paragraph that should be deleted by the engine.
        public string Empty { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Create a blank document and a builder to insert content.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add some static text.
            builder.Writeln("Header");

            // Insert a LINQ Reporting tag that resolves to an empty string.
            builder.Writeln("<<[model.Empty]>>");

            // Add more static text after the empty paragraph.
            builder.Writeln("Footer");

            // Prepare the data source.
            ReportModel model = new ReportModel();

            // Configure the reporting engine to remove empty paragraphs.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report using the model; the root name must match the tag prefix.
            engine.BuildReport(doc, model, "model");

            // Save the resulting document.
            const string outputPath = "ReportWithRemovedEmptyParagraphs.docx";
            doc.Save(outputPath);

            // Inform the user (optional, no input required).
            Console.WriteLine($"Report generated and saved to '{outputPath}'.");
        }
    }
}
