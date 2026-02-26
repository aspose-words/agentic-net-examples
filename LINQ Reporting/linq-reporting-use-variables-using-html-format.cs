using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define a document variable that holds HTML markup.
        string html = "<b>Bold Text</b><br/>Line 2";
        doc.Variables.Add("HtmlContent", html);

        // Insert a LINQ Reporting placeholder.
        // The ':html' suffix tells the engine to treat the replacement as HTML.
        builder.Writeln("<<[HtmlContent]:html>>");

        // Build the report. No external data source is needed; we only want the engine
        // to resolve the document variable.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, null, "");

        // Save the generated document.
        doc.Save("ReportWithHtmlVariable.docx");
    }
}
