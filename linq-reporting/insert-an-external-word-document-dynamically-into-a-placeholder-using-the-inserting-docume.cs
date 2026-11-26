using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

namespace InsertExternalDocumentExample
{
    // Model class used by the LINQ Reporting engine.
    public class ReportModel
    {
        // Path to the external Word document to be inserted.
        public string ExternalDocPath { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some Aspose.Words operations).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Define file names.
            const string externalDocFile = "External.docx";
            const string templateFile = "Template.docx";
            const string outputFile = "Output.docx";

            // -----------------------------------------------------------------
            // Step 1: Create the external Word document that will be inserted.
            // -----------------------------------------------------------------
            var externalDoc = new Document();
            var extBuilder = new DocumentBuilder(externalDoc);
            extBuilder.Writeln("This is the content of the external document.");
            externalDoc.Save(externalDocFile);

            // ---------------------------------------------------------------
            // Step 2: Create the template document with a placeholder tag.
            // ---------------------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);
            builder.Writeln("Report Start");
            // The <<doc>> tag inserts the external document at runtime.
            builder.Writeln("<<doc [model.ExternalDocPath]>>");
            builder.Writeln("Report End");
            templateDoc.Save(templateFile);

            // ---------------------------------------------------------------
            // Step 3: Load the template and prepare the model.
            // ---------------------------------------------------------------
            var loadedTemplate = new Document(templateFile);
            var model = new ReportModel
            {
                ExternalDocPath = Path.GetFullPath(externalDocFile)
            };

            // ---------------------------------------------------------------
            // Step 4: Build the report using the LINQ Reporting engine.
            // ---------------------------------------------------------------
            var engine = new ReportingEngine();
            engine.BuildReport(loadedTemplate, model, "model");

            // ---------------------------------------------------------------
            // Step 5: Save the generated report.
            // ---------------------------------------------------------------
            loadedTemplate.Save(outputFile);
        }
    }
}
