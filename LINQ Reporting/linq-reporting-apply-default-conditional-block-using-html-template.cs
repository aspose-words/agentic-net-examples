using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Load the HTML template. HtmlLoadOptions can be customized if needed.
        string templatePath = "Template.html";
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();
        Document doc = new Document(templatePath, loadOptions);

        // Prepare a simple data source that matches the template's placeholders.
        var data = new
        {
            persons = new[]
            {
                new { Name = "John Doe", Age = 30 },
                new { Name = "Jane Smith", Age = 15 }
            }
        };

        // Configure the ReportingEngine.
        // AllowMissingMembers lets the engine replace missing members with the MissingMemberMessage.
        // RemoveEmptyParagraphs cleans up paragraphs that become empty after processing.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers | ReportBuildOptions.RemoveEmptyParagraphs,
            MissingMemberMessage = "N/A"
        };

        // Build the report. The data source is exposed to the template as "data".
        engine.BuildReport(doc, data, "data");

        // Save the resulting document.
        doc.Save("Report.docx");
    }
}
