using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Product
{
    public int Id { get; set; }
    public string Name { get; set; } = "";
    public decimal Price { get; set; }
}

public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for Aspose.Words)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data
        var model = new ReportModel
        {
            Products = new List<Product>
            {
                new Product { Id = 1, Name = "Apple", Price = 0.99m },
                new Product { Id = 2, Name = "Banana", Price = 0.59m },
                new Product { Id = 3, Name = "Cherry", Price = 2.49m }
            }
        };

        // Create template document
        var template = new Document();
        var builder = new DocumentBuilder(template);

        builder.Writeln("Product Report");
        builder.Writeln();

        // Header row (outside foreach)
        var headerTable = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Id");
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Price");
        builder.EndRow();
        builder.EndTable();

        // Begin foreach over Products collection
        builder.Writeln("<<foreach [p in Products]>>");

        // Table row for each product
        var dataTable = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("<<[p.Id]>>");
        builder.InsertCell();
        builder.Writeln("<<[p.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[p.Price]>>");
        builder.EndRow();
        builder.EndTable();

        // End foreach
        builder.Writeln("<</foreach>>");

        // Save the template
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for report generation
        var reportDoc = new Document(templatePath);

        // Build the report
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report
        const string reportPath = "Report.docx";
        reportDoc.Save(reportPath);
    }
}
