using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

namespace LinqReportingInlineErrorExample
{
    public class Model
    {
        public string Name { get; set; } = "John Doe";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code pages provider required by Aspose.Words.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Create a template document with a correct tag, an incorrect tag, and the <<error>> placeholder.
            var templatePath = "template.docx";
            var builder = new DocumentBuilder();
            builder.Writeln("Customer: <<[model.Name]>>");
            builder.Writeln("Invalid expression: <<[model.NonExistent]>>");
            builder.Writeln("<<error>>");
            builder.Document.Save(templatePath);

            // Load the template for reporting.
            var doc = new Document(templatePath);

            // Prepare the data model.
            var model = new Model();

            // Configure the reporting engine to inline error messages.
            var engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            // Build the report.
            bool success = engine.BuildReport(doc, model, "model");

            // Save the generated report.
            var outputPath = "report.docx";
            doc.Save(outputPath);

            // Output the result.
            Console.WriteLine($"Report generation success: {success}");
            Console.WriteLine($"Report saved to: {Path.GetFullPath(outputPath)}");
        }
    }
}
