using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsExample
{
    // Simple data class that will be used as a data source for the report.
    public class Product
    {
        public string Name { get; set; }
        public decimal Price { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Prepare an HTML string that contains Aspose.Words reporting tags.
            //    The <<[ds.Name]>> and <<[ds.Price]:currency>> tags will be replaced by the engine.
            string htmlTemplate = @"
                <html>
                <body>
                    <h1>Product List</h1>
                    <<foreach [ds]>>
                        <p>Product: <<[ds.Name]>> – Price: <<[ds.Price]:currency>></p>
                    <<endforeach>>
                </body>
                </html>";

            // 2. Load the HTML string into a Document object.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertHtml(htmlTemplate);

            // 3. Create a data source using LINQ.
            //    Here we start from a simple collection and project it to a list of Product objects.
            var sourceData = new[]
            {
                new { Id = 1, Title = "Apple", Cost = 0.99m },
                new { Id = 2, Title = "Banana", Cost = 0.59m },
                new { Id = 3, Title = "Cherry", Cost = 2.49m }
            }
            .Select(item => new Product
            {
                Name = item.Title,
                Price = item.Cost
            })
            .ToList();

            // 4. Build the report using the ReportingEngine.
            //    The data source name "ds" matches the name used in the template tags.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, sourceData, "ds");

            // 5. Save the resulting document.
            doc.Save("ProductReport.docx");
        }
    }
}
