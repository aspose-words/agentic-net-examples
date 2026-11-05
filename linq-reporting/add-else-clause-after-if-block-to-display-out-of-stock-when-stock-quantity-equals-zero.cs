using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Product
{
    public string Name { get; set; } = "";
    public int Stock { get; set; }
}

public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        // Create template document.
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Write a heading.
        builder.Writeln("Product Stock Report");
        builder.Writeln();

        // Begin foreach over Products.
        builder.Writeln("<<foreach [p in Products]>>");
        builder.Writeln("Product: <<[p.Name]>>");
        // If stock > 0, show quantity, else show 'Out of stock'.
        builder.Writeln("<<if [p.Stock > 0]>>Quantity: <<[p.Stock]>> in stock<</if>>");
        builder.Writeln("<<if [p.Stock == 0]>>Out of stock<</if>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Prepare sample data.
        ReportModel model = new()
        {
            Products = new()
            {
                new Product { Name = "Apple", Stock = 15 },
                new Product { Name = "Banana", Stock = 0 },
                new Product { Name = "Cherry", Stock = 7 }
            }
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string resultPath = Path.Combine(outputDir, "report.docx");
        doc.Save(resultPath);
    }
}
