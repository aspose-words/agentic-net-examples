using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the Word template that contains LINQ Reporting conditional tags.
        // Example template content:
        //   <<if [ds.ShowSection]>>
        //   This section is shown.
        //   <<else>>
        //   This section is hidden.
        //   <<endif>>
        Document doc = new Document("Template.docx");

        // Prepare the data source for the template.
        var data = new ReportData
        {
            ShowSection = false,               // Triggers the else block.
            Title = "Quarterly Report"
        };

        // Create and configure the ReportingEngine.
        ReportingEngine engine = new ReportingEngine
        {
            // Allow missing members to be replaced with a custom message.
            Options = ReportBuildOptions.AllowMissingMembers,
            MissingMemberMessage = "N/A"
        };

        // Build the report, exposing the data source under the name "ds".
        engine.BuildReport(doc, data, "ds");

        // Save the populated document as PDF.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        doc.Save("Result.pdf", pdfOptions);
    }

    // Simple data class used as the data source for the template.
    public class ReportData
    {
        public bool ShowSection { get; set; }
        public string Title { get; set; }
    }
}
