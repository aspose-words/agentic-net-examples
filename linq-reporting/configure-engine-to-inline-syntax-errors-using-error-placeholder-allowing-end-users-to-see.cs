using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingInlineErrors
{
    // Simple data model used by the template.
    public class ReportModel
    {
        // Initialize to avoid nullable warnings.
        public string Name { get; set; } = "John Doe";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for Aspose.Words on .NET Core).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for the template and the generated report.
            string templatePath = "Template.docx";
            string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create a Word template with LINQ Reporting tags.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Correct tag – will be replaced with the model's Name value.
            builder.Writeln("Customer Name: <<[model.Name]>>");

            // Missing member – the engine will inline an error message here.
            builder.Writeln("Missing Property: <<[model.MissingProperty]>>");

            // Syntax error – malformed tag (no closing >>). The engine will also inline an error.
            builder.Writeln("Syntax Error Tag: <<[model.Name]");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);
            ReportModel model = new ReportModel();

            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            // BuildReport returns true when parsing succeeds (i.e., when InlineErrorMessages is set).
            bool success = engine.BuildReport(doc, model, "model");

            // Save the generated report.
            doc.Save(reportPath);

            // Output the result to the console.
            Console.WriteLine($"Report generation success flag: {success}");
            Console.WriteLine($"Report saved to: {Path.GetFullPath(reportPath)}");
        }
    }
}
