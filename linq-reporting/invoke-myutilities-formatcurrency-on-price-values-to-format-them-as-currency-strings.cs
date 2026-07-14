using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Utility class with a static method to format a decimal as a currency string.
    public static class MyUtilities
    {
        public static string FormatCurrency(decimal value)
        {
            // Use the current culture's currency format.
            return value.ToString("C");
        }
    }

    // Data model used by the LINQ Reporting engine.
    public class Product
    {
        // Sample price property.
        public decimal Price { get; set; } = 0m;
    }

    public class ReportModel
    {
        // Collection of products that will be iterated in the template.
        public List<Product> Items { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a blank document and a builder to construct the template.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // 2. Insert a foreach block that iterates over the Items collection.
            builder.Writeln("<<foreach [item in Items]>>");
            // Inside the loop, call the custom static method to format the price.
            builder.Writeln("Price: <<[MyUtilities.FormatCurrency(item.Price)]>>");
            builder.Writeln("<</foreach>>");

            // 3. Prepare sample data.
            ReportModel model = new ReportModel
            {
                Items = new List<Product>
                {
                    new Product { Price = 19.99m },
                    new Product { Price = 5.5m },
                    new Product { Price = 1234.56m }
                }
            };

            // 4. Configure the reporting engine.
            ReportingEngine engine = new ReportingEngine();
            // Register the utility class so its static members can be used in the template.
            engine.KnownTypes.Add(typeof(MyUtilities));

            // 5. Build the report using the model as the root object named "model".
            engine.BuildReport(doc, model, "model");

            // 6. Save the generated document.
            doc.Save("ReportOutput.docx");
        }
    }
}
