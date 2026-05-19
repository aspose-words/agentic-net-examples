using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // URL to be used for the hyperlink.
    public string Url { get; set; } = "https://www.example.com";

    // Text that will be displayed for the hyperlink.
    public string LinkText { get; set; } = "Visit Example";
}

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a LINQ Reporting link tag.
        // The tag will be replaced with a hyperlink whose URL and display text come from the data source.
        builder.Writeln("<<link [model.Url] [model.LinkText]>>");

        // Prepare the data source.
        ReportModel model = new ReportModel();

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        doc.Save("HyperlinkReport.docx");
    }
}
