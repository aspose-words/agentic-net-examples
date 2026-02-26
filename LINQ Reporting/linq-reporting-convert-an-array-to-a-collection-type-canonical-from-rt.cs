using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Loading;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model used in the template.
    public class Product
    {
        public string Name { get; set; }
        public double Price { get; set; }
    }

    // Wrapper class that contains an array of products.
    public class ReportData
    {
        public Product[] Products { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the RTF template. The template contains a LINQ Reporting tag such as:
            // <<foreach [ds.Products]>><<[Name]>> - $<<[Price]:currency>><</foreach>>
            // The RtfLoadOptions can be customized if needed.
            var loadOptions = new RtfLoadOptions();
            Document template = new Document("Template.rtf", loadOptions);

            // Prepare the data source with an array of Product objects.
            var data = new ReportData
            {
                Products = new[]
                {
                    new Product { Name = "Apple",  Price = 0.99 },
                    new Product { Name = "Banana", Price = 0.59 },
                    new Product { Name = "Cherry", Price = 2.49 }
                }
            };

            // Build the report using the LINQ Reporting engine.
            var engine = new ReportingEngine();
            // The data source name "ds" matches the name used in the template tags.
            engine.BuildReport(template, data, "ds");

            // Save the populated document.
            template.Save("Report.docx");

            // Example of converting the document's SectionCollection to an array (canonical collection type).
            Section[] sectionsArray = template.Sections.ToArray();

            // Output the number of sections to verify the conversion.
            Console.WriteLine($"Document contains {sectionsArray.Length} section(s).");
        }
    }
}
