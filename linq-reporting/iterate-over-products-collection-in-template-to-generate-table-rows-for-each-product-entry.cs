using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;   // Required for Table type

public class Product
{
    public string Name { get; set; } = "";
    public double Price { get; set; }
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
        // 1. Create the template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Begin foreach loop over the Products collection.
        builder.Writeln("<<foreach [p in Products]>>");

        // Table header.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Price");
        builder.InsertCell();
        builder.Writeln("Quantity");
        builder.EndRow();

        // Table row that will be repeated for each product.
        builder.InsertCell();
        builder.Writeln("<<[p.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[p.Price]>>");
        builder.InsertCell();
        builder.Writeln("<<[p.Quantity]>>");
        builder.EndRow();

        // End of the table and foreach block.
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // 2. Load the template and build the report.
        Document report = new Document(templatePath);

        // Sample data.
        ReportModel model = new ReportModel
        {
            Products = new List<Product>
            {
                new Product { Name = "Apple",  Price = 0.5, Quantity = 10 },
                new Product { Name = "Banana", Price = 0.3, Quantity = 15 },
                new Product { Name = "Orange", Price = 0.8, Quantity = 8 }
            }
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(report, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        report.Save(outputPath);
    }
}
