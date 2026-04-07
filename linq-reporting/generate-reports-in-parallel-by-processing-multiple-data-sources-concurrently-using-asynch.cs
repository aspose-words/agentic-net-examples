using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

// Ensure code page support for Aspose.Words if needed.
Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

public class LinqReportingDemo
{
    public static async Task Main(string[] args)
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create template for persons.
        // -----------------------------------------------------------------
        string personsTemplatePath = Path.Combine(outputDir, "persons_template.docx");
        CreatePersonsTemplate(personsTemplatePath);

        // -----------------------------------------------------------------
        // 2. Create template for products.
        // -----------------------------------------------------------------
        string productsTemplatePath = Path.Combine(outputDir, "products_template.docx");
        CreateProductsTemplate(productsTemplatePath);

        // -----------------------------------------------------------------
        // 3. Prepare sample data.
        // -----------------------------------------------------------------
        var personModel = new PersonModel
        {
            Persons = new List<Person>
            {
                new Person { Name = "Alice Johnson", Age = 30 },
                new Person { Name = "Bob Smith", Age = 45 },
                new Person { Name = "Charlie Davis", Age = 28 }
            }
        };

        var productModel = new ProductModel
        {
            Products = new List<Product>
            {
                new Product { Name = "Laptop", Price = 1299.99m },
                new Product { Name = "Smartphone", Price = 799.50m },
                new Product { Name = "Headphones", Price = 199.95m }
            }
        };

        // -----------------------------------------------------------------
        // 4. Generate reports in parallel using asynchronous rendering.
        // -----------------------------------------------------------------
        var tasks = new List<Task>
        {
            GenerateReportAsync(
                personsTemplatePath,
                personModel,
                "model",
                Path.Combine(outputDir, "persons_report.docx")),

            GenerateReportAsync(
                productsTemplatePath,
                productModel,
                "model",
                Path.Combine(outputDir, "products_report.docx"))
        };

        await Task.WhenAll(tasks);
    }

    // Asynchronously builds a report from a template and saves the result.
    private static async Task GenerateReportAsync(string templatePath, object model, string rootName, string outputPath)
    {
        // Load the template document.
        var doc = new Document(templatePath);

        // Configure the reporting engine.
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };

        // Build the report on a background thread to avoid blocking.
        await Task.Run(() => engine.BuildReport(doc, model, rootName));

        // Save the generated report.
        doc.Save(outputPath);
    }

    // Creates a simple persons report template with LINQ Reporting tags.
    private static void CreatePersonsTemplate(string filePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Persons Report");
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("- <<[p.Name]>> (Age: <<[p.Age]>>)");
        builder.Writeln("<</foreach>>");

        doc.Save(filePath);
    }

    // Creates a simple products report template with LINQ Reporting tags.
    private static void CreateProductsTemplate(string filePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Products Report");
        builder.Writeln("<<foreach [pr in Products]>>");
        builder.Writeln("- <<[pr.Name]>>: $<<[pr.Price]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(filePath);
    }
}

// ---------------------------------------------------------------------
// Data model classes – all members are public and initialized.
// ---------------------------------------------------------------------
public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}

public class Product
{
    public string Name { get; set; } = string.Empty;
    public decimal Price { get; set; }
}

// Wrapper classes used as the root objects for the templates.
public class PersonModel
{
    public List<Person> Persons { get; set; } = new();
}

public class ProductModel
{
    public List<Product> Products { get; set; } = new();
}
