using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsXpsExport
{
    // Simple data source class used in the template.
    public class ReportData
    {
        public string Title { get; set; }
        public string Author { get; set; }
        public DateTime Created { get; set; }
        public decimal Amount { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the DOCX template that contains Aspose.Words reporting tags.
            const string templatePath = @"C:\Templates\ReportTemplate.docx";

            // Path where the resulting XPS file will be saved.
            const string outputPath = @"C:\Output\ReportResult.xps";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Prepare the data source that will be merged into the template.
            var data = new ReportData
            {
                Title = "Quarterly Sales Report",
                Author = "John Doe",
                Created = DateTime.Now,
                Amount = 123456.78m
            };

            // Populate the template using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, data, "data"); // "data" is the name used in the template tags.

            // Create XpsSaveOptions to control XPS output (optional customizations can be set here).
            XpsSaveOptions xpsOptions = new XpsSaveOptions();

            // Save the populated document as XPS using the save overload that accepts SaveOptions.
            doc.Save(outputPath, xpsOptions);

            Console.WriteLine("Document successfully exported to XPS.");
        }
    }
}
