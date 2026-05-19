using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingInlineErrorDemo
{
    // Simple data model used by the template.
    public class ReportModel
    {
        // Initialise to avoid nullable warnings.
        public string Name { get; set; } = "John Doe";
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a blank document and insert a valid LINQ Reporting tag.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            // The tag references the Name property of the root object "model".
            builder.Writeln("<<[model.Name]>>");

            // 2. Prepare the data source.
            ReportModel model = new ReportModel();

            // 3. Configure the ReportingEngine to inline error messages.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            // 4. Build the report and capture the success flag.
            bool success = engine.BuildReport(doc, model, "model");

            // 5. Save the generated report.
            const string outputPath = "ReportWithInlineErrorMessages.docx";
            doc.Save(outputPath);

            // 6. Output the success flag.
            Console.WriteLine($"Report build success: {success}");
            Console.WriteLine($"Report saved to: {outputPath}");
        }
    }
}
