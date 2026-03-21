using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace InsertExternalDocumentExample
{
    // Simple data source class that will be referenced from the template.
    // The template should contain a tag like <<doc [src.Document]>> where "src" is the name we give to the data source.
    public class DocumentSource
    {
        public Document Document { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Determine the working directory (the folder where the executable runs).
            string workDir = AppDomain.CurrentDomain.BaseDirectory;

            // Paths for the template and the external document.
            string templatePath = Path.Combine(workDir, "Template.docx");
            string externalPath = Path.Combine(workDir, "External.docx");
            string resultPath   = Path.Combine(workDir, "Result.docx");

            // Ensure the template exists. If not, create a minimal one with the required insert tag.
            if (!File.Exists(templatePath))
            {
                Document templateDoc = new Document();
                DocumentBuilder builder = new DocumentBuilder(templateDoc);
                builder.Writeln("This is the main template.");
                // Insert the reporting tag that will be replaced with the external document.
                builder.Writeln("<<doc [src.Document]>>");
                templateDoc.Save(templatePath);
            }

            // Ensure the external document exists. If not, create a simple document.
            if (!File.Exists(externalPath))
            {
                Document externalDocTmp = new Document();
                DocumentBuilder builderTmp = new DocumentBuilder(externalDocTmp);
                builderTmp.Writeln("This is the content of the external document.");
                externalDocTmp.Save(externalPath);
            }

            // Load the template that contains the insert tag.
            Document template = new Document(templatePath);

            // Load the external Word document that we want to insert.
            Document externalDoc = new Document(externalPath);

            // Prepare the data source instance.
            var src = new DocumentSource { Document = externalDoc };

            // Use the ReportingEngine to process the template.
            // The second parameter is the name that will be used inside the tag (src in this case).
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, src, "src");

            // Save the resulting document.
            template.Save(resultPath);

            Console.WriteLine($"Report generated successfully: {resultPath}");
        }
    }
}
