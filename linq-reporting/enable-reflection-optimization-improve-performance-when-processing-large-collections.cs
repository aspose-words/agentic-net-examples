using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReflectionOptimizationDemo
{
    class Program
    {
        static void Main()
        {
            // Enable reflection optimization.
            // This makes the ReportingEngine generate dynamic proxy classes for
            // custom type members accessed via reflection, improving performance
            // when processing large collections.
            ReportingEngine.UseReflectionOptimization = true;

            // Load a template document that contains LINQ Reporting Engine tags.
            // Replace the path with the actual location of your template file.
            Document template = new Document(@"C:\Templates\ReportTemplate.docx");

            // Prepare a data source. For demonstration we use an anonymous object
            // with a collection of items. In a real scenario this could be any
            // object, DataTable, IEnumerable, etc.
            var data = new
            {
                Title = "Sales Report",
                Date = DateTime.Today,
                Items = new[]
                {
                    new { Product = "Apple",  Quantity = 120, Price = 0.5 },
                    new { Product = "Banana", Quantity = 85,  Price = 0.3 },
                    new { Product = "Orange", Quantity = 60,  Price = 0.4 }
                }
            };

            // Create the ReportingEngine instance.
            ReportingEngine engine = new ReportingEngine();

            // Build the report by merging the template with the data source.
            engine.BuildReport(template, data);

            // Save the generated report.
            // Replace the path with the desired output location.
            template.Save(@"C:\Output\SalesReport.docx");
        }
    }
}
