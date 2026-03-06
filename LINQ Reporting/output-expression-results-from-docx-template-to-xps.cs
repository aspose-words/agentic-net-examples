using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsXpsExport
{
    // Sample data source that will be used in the DOCX template.
    public class ReportData
    {
        public string Title { get; set; }
        public DateTime GeneratedOn { get; set; }
        public decimal Amount { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the DOCX template that contains Aspose.Words reporting tags.
            Document template = new Document("Template.docx");

            // Prepare the data that will be merged into the template.
            var data = new ReportData
            {
                Title = "Quarterly Sales Report",
                GeneratedOn = DateTime.Now,
                Amount = 123456.78m
            };

            // Populate the template using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, data, "ds"); // "ds" is the name used in the template.

            // Create XPS save options (default constructor is sufficient for most scenarios).
            XpsSaveOptions xpsOptions = new XpsSaveOptions();

            // Save the populated document as XPS.
            template.Save("ReportOutput.xps", xpsOptions);
        }
    }
}
