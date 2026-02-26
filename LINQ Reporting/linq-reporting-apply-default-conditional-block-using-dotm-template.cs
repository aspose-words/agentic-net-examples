using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOTM template that contains a conditional block, e.g.:
        // <<if [report.ShowDetails]>><<[report.Details]>><<else>>No details available.<</else>><</if>>
        Document doc = new Document("Template.dotm");

        // Prepare the data source that will be referenced in the template.
        var data = new ReportData
        {
            Title = "Quarterly Report",
            ShowDetails = false,               // Condition for the <<if>> block.
            Details = "Revenue increased by 15%."
        };

        // Create and configure the ReportingEngine.
        ReportingEngine engine = new ReportingEngine
        {
            // Remove paragraphs that become empty after the conditional block is evaluated.
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };

        // Build the report. The third argument is the name used in the template to reference the data source.
        engine.BuildReport(doc, data, "report");

        // Save the populated document.
        doc.Save("Result.docx");
    }

    // Simple POCO that matches the fields used in the template.
    public class ReportData
    {
        public string Title { get; set; }
        public bool ShowDetails { get; set; }
        public string Details { get; set; }
    }
}
