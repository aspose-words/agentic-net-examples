using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

public class Product
{
    public decimal Price { get; set; } = 0m;
    public string Name { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        var product = new Product
        {
            Name = "Sample Product",
            Price = 1234.56m
        };

        // Create a template document with a LINQ Reporting tag.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Product: <<[model.Name]>>");
        builder.Writeln("Price (formatted with culture): <<[model.Price]>>");
        // Save the template (optional, just for demonstration).
        const string templatePath = "template.docx";
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);

        // Configure a custom culture (e.g., French) for number formatting.
        var customCulture = new CultureInfo("fr-FR");
        System.Threading.Thread.CurrentThread.CurrentCulture = customCulture;
        System.Threading.Thread.CurrentThread.CurrentUICulture = customCulture;

        // Build the report using the custom culture.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // default options
        bool success = engine.BuildReport(reportDoc, product, "model");

        // Save the generated report.
        Directory.CreateDirectory("output");
        const string outputPath = "output/report.docx";
        reportDoc.Save(outputPath);

        // Indicate completion (no interactive input).
        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Saved to '{outputPath}'.");
    }
}
