using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // HTML snippet that will be inserted into the document.
    public string HtmlSnippet { get; set; } = "<p style='color:blue;'>This is <b>bold</b> and <i>italic</i> text.</p>";
}

public class Program
{
    public static void Main()
    {
        // Prepare the data source.
        var model = new ReportModel();

        // Create a blank template document.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert the LINQ Reporting tag that evaluates the HtmlSnippet property
        // and renders it as formatted HTML using the -html switch.
        builder.Writeln("<<[model.HtmlSnippet] -html>>");

        // Build the report by merging the template with the data source.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the resulting document.
        doc.Save("ReportWithHtml.docx");
    }
}
