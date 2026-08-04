using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Product
{
    public string Name { get; set; } = "";
    public double Quantity { get; set; }

    // Returns the quantity formatted as a simple fraction-like string (e.g., 1.5 -> "1/5").
    public string QuantityFormatted => Quantity.ToString("0/##");
}

public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Sample data.
        var model = new ReportModel
        {
            Products = new List<Product>
            {
                new Product { Name = "Apples", Quantity = 1.5 },
                new Product { Name = "Bananas", Quantity = 2.25 },
                new Product { Name = "Cherries", Quantity = 0.75 }
            }
        };

        // Build the template document.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        builder.Writeln("Products Report");
        builder.Writeln("<<foreach [p in Products]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        // Use the pre‑formatted quantity string.
        builder.Writeln("Qty: <<[p.QuantityFormatted]>>");
        builder.Writeln("<</foreach>>");

        // Save the template (optional, for inspection).
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template and generate the report.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        const string reportPath = "Report.docx";
        reportDoc.Save(reportPath);
    }
}
