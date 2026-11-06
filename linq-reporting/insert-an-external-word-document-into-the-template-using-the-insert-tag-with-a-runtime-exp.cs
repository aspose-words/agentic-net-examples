using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

namespace InsertExternalDocExample
{
    public class ReportModel
    {
        // Path to the external document to be inserted.
        public string ExternalDocPath { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for Aspose.Words).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare directories.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
            Directory.CreateDirectory(outputDir);

            // Create an external Word document.
            string externalDocPath = Path.Combine(outputDir, "External.docx");
            Document externalDoc = new Document();
            DocumentBuilder extBuilder = new DocumentBuilder(externalDoc);
            extBuilder.Writeln("This is the content of the external document.");
            externalDoc.Save(externalDocPath);

            // Create a template document with an insert tag that references the external document at runtime.
            string templatePath = Path.Combine(outputDir, "Template.docx");
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);
            builder.Writeln("Report Start");
            // Insert tag using a runtime expression.
            builder.Writeln("<<doc [model.ExternalDocPath]>>");
            builder.Writeln("Report End");
            templateDoc.Save(templatePath);

            // Load the template for reporting.
            Document loadedTemplate = new Document(templatePath);

            // Prepare the data model.
            ReportModel model = new ReportModel
            {
                ExternalDocPath = externalDocPath
            };

            // Build the report.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(loadedTemplate, model, "model");

            // Save the final report.
            string resultPath = Path.Combine(outputDir, "ReportOutput.docx");
            loadedTemplate.Save(resultPath);
        }
    }
}
