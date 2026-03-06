using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsExample
{
    // Simple data source class with properties used in the template.
    public class ReportData
    {
        public string Title { get; set; }
        public decimal Amount { get; set; }
        public DateTime Date { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the PDF template document.
            // Aspose.Words can open PDF files as a Document object.
            Document doc = new Document("Template.pdf");

            // Prepare the data source that will be merged into the template.
            ReportData data = new ReportData
            {
                Title = "Quarterly Report",
                Amount = 12345.67m,
                Date = DateTime.Today
            };

            // Build the report by populating the template with the data source.
            ReportingEngine engine = new ReportingEngine();
            // The second parameter is the data source object; the third parameter is the name
            // used to reference the data source inside the template (e.g., <<[data.Title]>>).
            engine.BuildReport(doc, data, "data");

            // Configure HTML Fixed save options.
            HtmlFixedSaveOptions htmlOptions = new HtmlFixedSaveOptions
            {
                // Export form fields as interactive HTML input elements.
                ExportFormFields = true,
                // Optional: keep the output tidy.
                PrettyFormat = true,
                // Optional: embed CSS directly into the HTML.
                ExportEmbeddedCss = true
            };

            // Save the populated document as fixed HTML.
            doc.Save("Report.html", htmlOptions);
        }
    }
}
