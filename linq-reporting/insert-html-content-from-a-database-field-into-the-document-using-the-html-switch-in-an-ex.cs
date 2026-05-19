using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    // Simple data model representing a record from a database.
    public class ReportData
    {
        // HTML fragment that would normally come from a DB field.
        public string HtmlContent { get; set; } = "<h2 style='color:Blue;'>Hello from the database!</h2>"
                                                + "<p>This paragraph is <b>bold</b> and <i>italic</i>.</p>";
    }

    public static void Main()
    {
        // 1. Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert the LINQ Reporting tag that will render the HTML content.
        // The "-html" switch tells the engine to treat the expression result as HTML.
        builder.Writeln("<<[model.HtmlContent] -html>>");

        // 2. Prepare the data source (simulating a DB record).
        ReportData data = new ReportData();

        // 3. Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        // No special options are required for this simple scenario.
        engine.BuildReport(template, data, "model");

        // 4. Save the generated document.
        template.Save("ReportWithHtml.docx");
    }
}
