using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an HTML fragment that contains a template expression.
        // The expression references a member that does not exist in the data source.
        string html = "<p>Customer: <<[customer.Name]>></p>";
        builder.InsertHtml(html);

        // Configure the ReportingEngine to allow missing members.
        // Set a custom message that will be printed for any missing member.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers,
            MissingMemberMessage = "N/A"
        };

        // Build the report using an empty DataSet (no data source).
        engine.BuildReport(doc, new DataSet(), "");

        // Save the resulting document as HTML.
        doc.Save("Output.html", SaveFormat.Html);
    }
}
