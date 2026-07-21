using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Sample data model.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    public class ReportProgram
    {
        public static void Main()
        {
            // Required for code page support.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample data.
            var model = new Person { Name = "Alice", Age = 30 };

            // Create a template document in memory.
            var doc = new Document();
            var builder = new DocumentBuilder(doc);
            // Insert a LINQ Reporting tag that references the model.
            builder.Writeln("<<[model.Name]>> is <<[model.Age]>> years old.");

            // Configure the reporting engine to inline error messages.
            var engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            // Build the report and capture the success flag.
            bool success = engine.BuildReport(doc, model, "model");

            // Output the result.
            Console.WriteLine($"Report build success: {success}");

            // Save the generated document.
            doc.Save("ReportOutput.docx");
        }
    }
}
