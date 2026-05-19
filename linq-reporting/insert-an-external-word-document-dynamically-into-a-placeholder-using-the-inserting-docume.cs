using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Wrapper class for the data source used by the LINQ Reporting engine.
    public class ReportModel
    {
        // The external document that will be inserted into the template.
        public Document Document { get; set; } = new Document();
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // Step 1: Create the external Word document that will be inserted.
            // -----------------------------------------------------------------
            Document externalDoc = new Document();
            DocumentBuilder externalBuilder = new DocumentBuilder(externalDoc);
            externalBuilder.Writeln("This is the content of the external document.");
            externalBuilder.Writeln("It will be inserted dynamically into the template.");
            const string externalPath = "ExternalDocument.docx";
            externalDoc.Save(externalPath);

            // -----------------------------------------------------------------
            // Step 2: Create the template document with a placeholder tag.
            // The tag <<doc [src.Document]>> tells the reporting engine to insert
            // the Document object referenced by src.Document.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);
            templateBuilder.Writeln("=== Report Start ===");
            templateBuilder.Writeln("<<doc [src.Document]>>"); // Placeholder for the external document.
            templateBuilder.Writeln("=== Report End ===");
            const string templatePath = "TemplateDocument.docx";
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // Step 3: Load the saved template and the external document.
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templatePath);
            Document loadedExternal = new Document(externalPath);

            // -----------------------------------------------------------------
            // Step 4: Prepare the data model for the reporting engine.
            // The root name "src" must match the name used in the template tag.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel
            {
                Document = loadedExternal
            };

            // -----------------------------------------------------------------
            // Step 5: Build the report using the LINQ Reporting engine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(loadedTemplate, model, "src");

            // -----------------------------------------------------------------
            // Step 6: Save the final document.
            // -----------------------------------------------------------------
            const string outputPath = "ReportWithInsertedDocument.docx";
            loadedTemplate.Save(outputPath);
        }
    }
}
