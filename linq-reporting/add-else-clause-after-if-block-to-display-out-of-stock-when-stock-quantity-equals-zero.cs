using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Product
{
    public string Name { get; set; } = "";
    public int Quantity { get; set; }
}

public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Create sample data.
        var model = new ReportModel
        {
            Products = new List<Product>
            {
                new Product { Name = "Apple", Quantity = 5 },
                new Product { Name = "Banana", Quantity = 0 },
                new Product { Name = "Cherry", Quantity = 12 }
            }
        };

        // -----------------------------------------------------------------
        // Step 1: Build the template document programmatically.
        // -----------------------------------------------------------------
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Header.
        builder.Writeln("Product Stock Report");
        builder.Writeln();

        // Begin a foreach loop over the Products collection.
        builder.Writeln("<<foreach [p in Products]>>");

        // Product name.
        builder.Writeln("Product: <<[p.Name]>>");

        // If quantity > 0, show the quantity.
        builder.Writeln("<<if [p.Quantity > 0]>>Quantity: <<[p.Quantity]>> <</if>>");

        // Else (quantity == 0), show "Out of stock".
        builder.Writeln("<<if [p.Quantity == 0]>>Out of stock<</if>>");

        // End of foreach.
        builder.Writeln("<</foreach>>");

        // Save the template to a local file.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 2: Load the template and build the report.
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report using the model as the root data source.
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
