using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data
        var product = new Product
        {
            Name = "Widget",
            Quantity = 0 // Change to a positive number to see the quantity displayed
        };

        // Create template document with LINQ Reporting tags
        string templatePath = "template.docx";
        CreateTemplate(templatePath);

        // Load the template
        var doc = new Document(templatePath);

        // Build the report
        var engine = new ReportingEngine();
        engine.BuildReport(doc, product, "product");

        // Save the generated report
        string reportPath = "report.docx";
        doc.Save(reportPath);

        Console.WriteLine($"Report generated: {Path.GetFullPath(reportPath)}");
    }

    private static void CreateTemplate(string path)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Product: <<[product.Name]>>");
        // Inline if‑else: display quantity when >0, otherwise show "Out of stock"
        builder.Writeln("Stock: <<if [product.Quantity > 0]>> <<[product.Quantity]>> <<else>>Out of stock<</if>>");

        doc.Save(path);
    }
}

public class Product
{
    public string Name { get; set; } = string.Empty;
    public int Quantity { get; set; }
}
