using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Data model that provides the external document to be inserted.
    public class ReportData
    {
        // The external document that will be inserted via the <<doc>> tag.
        public Document Document { get; set; }

        public ReportData(Document document)
        {
            Document = document ?? throw new ArgumentNullException(nameof(document));
        }
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create an external Word document that will be inserted later.
            // -----------------------------------------------------------------
            const string externalDocPath = "External.docx";
            Document externalDoc = new Document();
            DocumentBuilder externalBuilder = new DocumentBuilder(externalDoc);
            externalBuilder.Writeln("This is the content of the external document.");
            externalDoc.Save(externalDocPath);

            // ---------------------------------------------------------------
            // 2. Create a template document containing the <<doc>> tag.
            //    The tag uses a runtime expression that references the data source.
            // ---------------------------------------------------------------
            const string templatePath = "Template.docx";
            Document templateDoc = new Document();
            DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);
            templateBuilder.Writeln("=== Report Start ===");
            // The insert tag: <<doc [src.Document]>>
            // It will be replaced with the content of the external document at runtime.
            templateBuilder.Writeln("<<doc [src.Document]>>");
            templateBuilder.Writeln("=== Report End ===");
            templateDoc.Save(templatePath);

            // ---------------------------------------------------------------
            // 3. Load the template (optional – we already have it in memory).
            // ---------------------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // ---------------------------------------------------------------
            // 4. Prepare the data source that supplies the external document.
            // ---------------------------------------------------------------
            ReportData data = new ReportData(externalDoc);

            // ---------------------------------------------------------------
            // 5. Build the report using Aspose.Words LINQ Reporting Engine.
            // ---------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // No special options required.
            engine.BuildReport(loadedTemplate, data, "src");

            // ---------------------------------------------------------------
            // 6. Save the final document.
            // ---------------------------------------------------------------
            const string resultPath = "Result.docx";
            loadedTemplate.Save(resultPath);
        }
    }
}
