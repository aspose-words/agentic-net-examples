using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Simple data model used by the template.
    public class ReportModel
    {
        // Regular property with a value.
        public string Name { get; set; } = "John Doe";

        // Property that returns an empty string – the paragraph will become empty after processing.
        public string Empty { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a template that contains:
            // 1. A valid expression.
            // 2. An expression referencing a missing member (will generate an inline error).
            // 3. An expression that evaluates to an empty string (will produce an empty paragraph).
            builder.Writeln("Customer: <<[model.Name]>>");
            builder.Writeln("Missing member: <<[model.Missing]>>");
            builder.Writeln("Empty value: <<[model.Empty]>>");

            // Configure the reporting engine to:
            // - Remove paragraphs that become empty after tag processing.
            // - Inline any syntax errors directly into the output document.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.RemoveEmptyParagraphs | ReportBuildOptions.InlineErrorMessages
            };

            // Build the report using the model as the data source.
            // The third parameter ("model") matches the root name used in the template tags.
            bool success = engine.BuildReport(doc, new ReportModel(), "model");

            // Output the result of the build (true = no parsing errors, false = errors were inlined).
            Console.WriteLine($"Report build successful: {success}");

            // Save the generated document.
            const string outputPath = "output.docx";
            doc.Save(outputPath);
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
