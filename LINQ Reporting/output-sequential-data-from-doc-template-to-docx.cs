using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsExample
{
    // Simple data class used as a data source for the template.
    public class ReportData
    {
        public string Title { get; set; }
        public string Author { get; set; }
        public DateTime Date { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the DOC template. The constructor automatically detects the format.
            Document templateDoc = new Document("Template.doc");

            // Prepare data that will be merged into the template.
            ReportData data = new ReportData
            {
                Title = "Quarterly Sales Report",
                Author = "John Doe",
                Date = DateTime.Today
            };

            // Populate the template using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // The data source name ("data") is used inside the template as <<[data.Title]>>, etc.
            engine.BuildReport(templateDoc, data, "data");

            // Save the populated document as DOCX.
            templateDoc.Save("Result.docx");
        }
    }
}
