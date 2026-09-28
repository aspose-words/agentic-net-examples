using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingKnownTypesDemo
{
    // Sample external type 1
    public class Person
    {
        public string Name { get; set; } = "John Doe";
        public int Age { get; set; } = 30;
    }

    // Sample external type 2
    public class Product
    {
        public string Title { get; set; } = "Sample Product";
        public decimal Price { get; set; } = 99.99m;
    }

    // Root model that holds the external objects
    public class ReportModel
    {
        public Person Person { get; set; } = new();
        public Product Product { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Ensure the output directory exists
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // -----------------------------------------------------------------
            // 1. Create the template document with LINQ Reporting tags
            // -----------------------------------------------------------------
            string templatePath = Path.Combine(outputDir, "Template.docx");
            Document templateDoc = new();
            DocumentBuilder builder = new(templateDoc);

            builder.Writeln("Person Name: <<[model.Person.Name]>>");
            builder.Writeln("Person Age: <<[model.Person.Age]>>");
            builder.Writeln("Product Title: <<[model.Product.Title]>>");
            builder.Writeln("Product Price: <<[model.Product.Price]>>");

            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Enable reflection optimization (allows known types without reflection)
            // -----------------------------------------------------------------
            ReportingEngine.UseReflectionOptimization = true;

            // -----------------------------------------------------------------
            // 3. Load the template and build the report
            // -----------------------------------------------------------------
            Document reportDoc = new(templatePath);
            ReportModel model = new()
            {
                Person = new Person { Name = "Alice Smith", Age = 28 },
                Product = new Product { Title = "Aspose.Words Book", Price = 49.95m }
            };

            ReportingEngine engine = new();
            engine.BuildReport(reportDoc, model, "model");

            // -----------------------------------------------------------------
            // 4. Save the generated report
            // -----------------------------------------------------------------
            string reportPath = Path.Combine(outputDir, "Report.docx");
            reportDoc.Save(reportPath);

            // -----------------------------------------------------------------
            // 5. Verify the output by printing the document text to console
            // -----------------------------------------------------------------
            Console.WriteLine("=== Generated Report Text ===");
            Console.WriteLine(reportDoc.GetText());
        }
    }
}
