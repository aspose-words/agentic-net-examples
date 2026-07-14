using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

namespace MyApp.Models
{
    // Simple person model.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }
}

namespace MyApp.Entities
{
    // Simple product model.
    public class Product
    {
        public string Name { get; set; } = string.Empty;
        public decimal Price { get; set; }
    }
}

// Wrapper model that will be passed to the reporting engine.
public class ReportModel
{
    public List<MyApp.Models.Person> Persons { get; set; } = new();
    public List<MyApp.Entities.Product> Products { get; set; } = new();
}

// Marked as partial to match the test harness's partial declaration.
public partial class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a template document programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Persons section.
        builder.Writeln("Persons:");
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("- <<[p.Name]>> (Age: <<[p.Age]>>)");
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        // Products section.
        builder.Writeln("Products:");
        builder.Writeln("<<foreach [pr in Products]>>");
        builder.Writeln("- <<[pr.Name]>> : $<<[pr.Price]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        // Example of using a static member from a registered external type.
        builder.Writeln("Generated GUID: <<[Guid.NewGuid()]>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template back (simulating a real-world scenario).
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare sample data.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel
        {
            Persons = new List<MyApp.Models.Person>
            {
                new() { Name = "Alice", Age = 30 },
                new() { Name = "Bob", Age = 45 }
            },
            Products = new List<MyApp.Entities.Product>
            {
                new() { Name = "Laptop", Price = 999.99m },
                new() { Name = "Smartphone", Price = 499.50m }
            }
        };

        // -----------------------------------------------------------------
        // 4. Configure the ReportingEngine and register external types.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();

        // Register types from different namespaces so that the template can access them.
        engine.KnownTypes.Add(typeof(System.Guid));                     // System namespace.
        engine.KnownTypes.Add(typeof(MyApp.Models.Person));            // MyApp.Models namespace.
        engine.KnownTypes.Add(typeof(MyApp.Entities.Product));         // MyApp.Entities namespace.

        // -----------------------------------------------------------------
        // 5. Build the report.
        // -----------------------------------------------------------------
        // The root object name in the template is "model".
        engine.BuildReport(loadedTemplate, model, "model");

        // -----------------------------------------------------------------
        // 6. Save the generated report.
        // -----------------------------------------------------------------
        const string outputPath = "Report.docx";
        loadedTemplate.Save(outputPath);
    }
}
