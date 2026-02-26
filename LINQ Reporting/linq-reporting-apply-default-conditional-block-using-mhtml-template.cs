using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;
using Aspose.Words.Loading;

namespace AsposeWordsMhtmlReport
{
    // Simple data source class used by the LINQ Reporting Engine.
    public class ReportData
    {
        public string Title { get; set; }
        public bool ShowDetails { get; set; }
        public string Details { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the MHTML template. HtmlLoadOptions are used because the template is HTML/MHTML.
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            Document doc = new Document("Template.mht", loadOptions);

            // Prepare the data source that will be referenced in the template.
            ReportData data = new ReportData
            {
                Title = "Quarterly Report",
                ShowDetails = true,
                Details = "Revenue increased by 15% compared to the previous quarter."
            };

            // Create the ReportingEngine and configure it.
            ReportingEngine engine = new ReportingEngine
            {
                // Remove empty paragraphs that may appear after conditional blocks are evaluated.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report. The data source name ("ds") can be used inside the template.
            engine.BuildReport(doc, data, "ds");

            // Save the resulting document as MHTML.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml);
            doc.Save("Report.mht", saveOptions);
        }
    }
}
