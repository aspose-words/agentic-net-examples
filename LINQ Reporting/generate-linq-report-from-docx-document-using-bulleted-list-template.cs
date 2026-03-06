using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Lists;

namespace AsposeWordsLinqReport
{
    // Simple data model for the report.
    public class Product
    {
        public string Name { get; set; }
        public decimal Price { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a list of products that will be used as the LINQ data source.
            List<Product> products = new List<Product>
            {
                new Product { Name = "Apple",  Price = 0.99m },
                new Product { Name = "Banana", Price = 0.59m },
                new Product { Name = "Cherry", Price = 2.49m }
            };

            // 2. Create a blank Word document (lifecycle: create).
            Document doc = new Document();

            // 3. Build the template that contains a bulleted list with Reporting Engine tags.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Title.
            builder.Writeln("Product List:");

            // Apply a bulleted list style (ListTemplate.BulletDefault) to the following paragraphs.
            List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);
            builder.ListFormat.List = bulletList;

            // Reporting Engine tags:
            // <<foreach [Products]>> – iterate over the collection named "Products".
            // <<[Name]>> – insert the Name property of the current item.
            // <</foreach>> – end of the loop.
            builder.Writeln("<<foreach [Products]>>");
            builder.Writeln("<<[Name]>> - $<<[Price]>>");
            builder.Writeln("<</foreach>>");

            // Remove list formatting for any subsequent paragraphs.
            builder.ListFormat.RemoveNumbers();

            // 4. Populate the template using the ReportingEngine (lifecycle: load is implicit in BuildReport).
            ReportingEngine engine = new ReportingEngine();

            // The anonymous object provides a property "Products" that the template can reference.
            var dataSource = new { Products = products };

            // BuildReport parses the template and inserts the data.
            engine.BuildReport(doc, dataSource);

            // 5. Save the resulting document (lifecycle: save).
            doc.Save("LinqReport.docx");
        }
    }
}
