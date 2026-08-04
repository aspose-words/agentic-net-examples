using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsInsertDocDynamic
{
    public class Program
    {
        public static void Main()
        {
            // Create the external document that will be inserted.
            Document externalDoc = new Document();
            DocumentBuilder externalBuilder = new DocumentBuilder(externalDoc);
            externalBuilder.Writeln("This is content from the external document.");
            const string externalPath = "External.docx";
            externalDoc.Save(externalPath);

            // Create the template document containing a placeholder for the external document.
            Document templateDoc = new Document();
            DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);
            templateBuilder.Writeln("Report start");
            // Placeholder tag that inserts the document provided by the data source.
            templateBuilder.Writeln("<<doc [src.Document]>>");
            templateBuilder.Writeln("Report end");

            // Load the external document into the data model.
            Document loadedExternal = new Document(externalPath);
            ReportData data = new ReportData(loadedExternal);

            // Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(templateDoc, data, "src");

            // Save the final document with the external document inserted.
            const string outputPath = "ReportWithInsertedDoc.docx";
            templateDoc.Save(outputPath);
        }
    }

    // Wrapper class that exposes the external Document to the template.
    public class ReportData
    {
        public Document Document { get; set; }

        public ReportData(Document document)
        {
            Document = document;
        }
    }
}
