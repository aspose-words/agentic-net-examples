using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsExample
{
    // Simple data source class whose members will be referenced in the template.
    public class ReportData
    {
        public string Title { get; set; }
        public int Quantity { get; set; }
        public decimal Price { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the DOC template that contains expression tags like <<[ds.Title]>>.
            string templatePath = @"C:\Templates\ReportTemplate.docx";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Prepare the data source instance.
            ReportData data = new ReportData
            {
                Title = "Sample Product",
                Quantity = 42,
                Price = 123.45m
            };

            // Create the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Populate the template with data.
            // The third argument ("ds") is the name used in the template to reference the data source.
            engine.BuildReport(doc, data, "ds");

            // Save the populated document as DOCX.
            string outputPath = @"C:\Output\ReportResult.docx";
            doc.Save(outputPath, SaveFormat.Docx);
        }
    }
}
