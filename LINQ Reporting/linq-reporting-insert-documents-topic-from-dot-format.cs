using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class InsertDocumentViaReportingEngine
{
    static void Main()
    {
        // 1. Create a template document that contains a reporting tag.
        // The tag <<doc [src.Document]>> tells the ReportingEngine to insert the document
        // referenced by the data source named "src".
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Report start");
        builder.Writeln("<<doc [src.Document]>>"); // placeholder for the inserted document
        builder.Writeln("Report end");

        // 2. Create the document that will be inserted.
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("This is the content of the inserted document.");
        srcBuilder.Writeln("It can contain multiple paragraphs, tables, images, etc.");

        // 3. Build the report.
        // The ReportingEngine receives the template and an array of data sources.
        // The source document is passed as a data source with the name "src".
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, new object[] { sourceDoc }, new string[] { "src" });

        // 4. Save the resulting document.
        template.Save("ReportWithInsertedDocument.docx");
    }
}
