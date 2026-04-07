using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Custom type with an explicit conversion operator to int.
    public class MyNumber
    {
        public int Value { get; set; }

        public MyNumber(int value) => Value = value;

        // Explicit conversion to int.
        public static explicit operator int(MyNumber number) => number.Value;
    }

    // Data model used by the report.
    public class Product
    {
        public string Name { get; set; } = "";
        public MyNumber Price { get; set; } = new MyNumber(0);
    }

    // Wrapper class that will be passed as the root data source.
    public class ReportModel
    {
        public List<Product> Products { get; set; } = new();
    }

    class Program
    {
        static void Main()
        {
            // 1. Create the template document programmatically.
            var template = new Document();
            var builder = new DocumentBuilder(template);

            // Header.
            builder.Writeln("Product Report");
            builder.Writeln();

            // Begin a foreach loop over the Products collection.
            builder.Writeln("<<foreach [p in Products]>>");

            // Write product name.
            builder.Writeln("Name: <<[p.Name]>>");

            // Explicitly cast the custom type to int within the template expression.
            builder.Writeln("Price: <<[(int)p.Price]>>");

            // End the foreach loop.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // 2. Prepare sample data.
            var model = new ReportModel
            {
                Products = new List<Product>
                {
                    new Product { Name = "Apple",  Price = new MyNumber(150) },
                    new Product { Name = "Banana", Price = new MyNumber( 80) },
                    new Product { Name = "Cherry", Price = new MyNumber(200) }
                }
            };

            // 3. Load the template (demonstrating the load step).
            var doc = new Document(templatePath);

            // 4. Build the report using the ReportingEngine.
            var engine = new ReportingEngine();
            // No special options are required for this simple example.
            engine.BuildReport(doc, model, "model");

            // 5. Save the generated report.
            const string reportPath = "Report.docx";
            doc.Save(reportPath);
        }
    }
}
