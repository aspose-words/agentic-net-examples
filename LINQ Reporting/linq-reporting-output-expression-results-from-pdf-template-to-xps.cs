using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace LinqReportingToXps
{
    // Simple data model used in the template.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the PDF template that contains LINQ Reporting tags, e.g. <<[person.Name]>>.
            string pdfTemplatePath = @"C:\Templates\ReportTemplate.pdf";

            // Load the PDF template into an Aspose.Words Document.
            Document doc = new Document(pdfTemplatePath);

            // Prepare the data source. Any non‑dynamic .NET object can be used.
            var data = new Person { Name = "John Doe", Age = 42 };

            // Build the report using the ReportingEngine.
            // The third argument ("person") is the name used inside the template to reference the data source.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, data, "person");

            // Configure XPS save options (optional customizations can be set here).
            XpsSaveOptions xpsOptions = new XpsSaveOptions
            {
                // Example: enable output optimization to reduce file size.
                OptimizeOutput = true
            };

            // Save the populated document as XPS.
            string outputXpsPath = @"C:\Output\ReportResult.xps";
            doc.Save(outputXpsPath, xpsOptions);

            Console.WriteLine("Report generated and saved to XPS successfully.");
        }
    }
}
