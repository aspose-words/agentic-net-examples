using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        var product = new Product
        {
            Name = "SuperWidget",
            Price = 49.99m,
            Discount = 15m // Discount percentage greater than zero.
        };

        // Create a template document programmatically.
        var templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Product: <<[product.Name]>>");
        builder.Writeln("Price: $<<[product.Price]>>");
        builder.Writeln("<<if [product.Discount > 0]>>Discount: <<[product.Discount]>>%<</if>>");

        doc.Save(templatePath);

        // Load the template for reporting.
        var template = new Document(templatePath);

        // Build the report using LINQ Reporting Engine.
        var engine = new ReportingEngine();
        engine.BuildReport(template, product, "product");

        // Save the generated report.
        var outputPath = "report.docx";
        template.Save(outputPath);
    }
}

// Public data model class.
public class Product
{
    public string Name { get; set; } = "";
    public decimal Price { get; set; }
    public decimal Discount { get; set; }
}
