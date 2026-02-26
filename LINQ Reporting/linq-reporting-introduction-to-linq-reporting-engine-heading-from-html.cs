using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the heading as HTML.
        string html = "<h1>LINQ Reporting Introduction to LINQ Reporting Engine</h1>";
        builder.InsertHtml(html);

        // Use the ReportingEngine to process the document (no data source needed for static content).
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, new object());

        // Save the resulting document.
        doc.Save("LINQReportingHeading.docx");
    }
}
