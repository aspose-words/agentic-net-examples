using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingEscapeKeyword
{
    // Data model with a property named "new". The @ prefix escapes the C# keyword.
    public class Model
    {
        public string @new { get; set; } = "EscapedKeywordValue";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some Aspose.Words operations).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Create a temporary folder for the example files.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
            Directory.CreateDirectory(outputDir);

            // -----------------------------------------------------------------
            // Step 1: Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a simple paragraph that contains a LINQ Reporting tag.
            // The property name "new" is referenced without the @ prefix inside the expression.
            builder.Writeln("Value: <<[model.new]>>");

            // Save the template to disk.
            string templatePath = Path.Combine(outputDir, "template.docx");
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // Step 2: Load the template and build the report.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // Prepare the root data object.
            Model model = new Model();

            // Create the reporting engine and generate the report.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report.
            string resultPath = Path.Combine(outputDir, "result.docx");
            reportDoc.Save(resultPath);
        }
    }
}
