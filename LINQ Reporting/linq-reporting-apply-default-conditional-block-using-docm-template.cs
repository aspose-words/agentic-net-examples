using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Reporting; // for ReportBuildOptions

namespace AsposeWordsLinqReportingDemo
{
    // Simple data source class used in the template.
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
            // Load the DOCM template that contains a default conditional block,
            // e.g. <<if [ds.ShowDetails]>>...<<endif>>.
            Document template = new Document("Template.docm");

            // Prepare the data source instance.
            ReportData data = new ReportData
            {
                Title = "Quarterly Summary",
                ShowDetails = true,               // Controls the conditional block.
                Details = "Revenue increased by 12%."
            };

            // Configure the reporting engine.
            ReportingEngine engine = new ReportingEngine
            {
                // Remove paragraphs that become empty after the conditional block is evaluated.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report. The data source name ("ds") must match the name used in the template.
            engine.BuildReport(template, data, "ds");

            // Save the populated document.
            template.Save("Report.docx");
        }
    }
}
