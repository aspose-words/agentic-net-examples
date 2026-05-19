using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    public string Url { get; set; } = "";
    public string LinkText { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Create a blank document and a builder to insert content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a LINQ Reporting tag that creates a hyperlink.
        // The tag will be replaced with a clickable link during report generation.
        builder.Writeln("<<link [model.Url] [model.LinkText]>>");

        // Prepare the data source.
        ReportModel model = new ReportModel
        {
            Url = "https://www.example.com",
            LinkText = "Visit Example"
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        doc.Save("HyperlinkReport.docx");
    }
}
