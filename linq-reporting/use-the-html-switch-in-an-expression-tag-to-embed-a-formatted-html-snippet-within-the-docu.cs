using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Sample HTML snippet that will be inserted into the document.
    public string HtmlSnippet { get; set; } = "<p style='color:blue;'>Hello <b>World</b>!</p>";
}

public class Program
{
    public static void Main()
    {
        // Create a blank document and a builder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a LINQ Reporting tag that evaluates the HtmlSnippet property
        // and renders it as formatted HTML using the -html switch.
        builder.Writeln("<<[model.HtmlSnippet] -html>>");

        // Prepare the data source.
        ReportModel model = new ReportModel();

        // Build the report by merging the template with the data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the resulting document.
        doc.Save("ReportWithHtml.docx");
    }
}
