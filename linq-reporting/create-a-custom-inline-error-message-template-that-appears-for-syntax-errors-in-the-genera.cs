using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace InlineErrorMessageExample
{
    // Simple data model used as the root object for the report.
    public class ReportModel
    {
        // Initialize to avoid nullable warnings.
        public string Name { get; set; } = "John Doe";
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write a paragraph that contains a correct LINQ Reporting tag.
            builder.Writeln("Customer name: <<[model.Name]>>");

            // Write a paragraph that contains a malformed tag – this will trigger a syntax error.
            // The missing closing ">>" is intentional to demonstrate inline error messages.
            builder.Writeln("This line has a syntax error: <<[model.Age]");

            // Prepare the data source.
            ReportModel model = new ReportModel();

            // Configure the reporting engine to inline error messages.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            // Build the report. The returned flag indicates whether parsing succeeded.
            // With InlineErrorMessages enabled, the method returns false when there are syntax errors,
            // but the document will contain the error messages inline.
            bool success = engine.BuildReport(doc, model, "model");

            // Output the result of the build operation to the console.
            Console.WriteLine($"Report build success: {success}");

            // Save the generated document. The inline error messages will be visible in the output file.
            string outputPath = "InlineErrorReport.docx";
            doc.Save(outputPath);
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
