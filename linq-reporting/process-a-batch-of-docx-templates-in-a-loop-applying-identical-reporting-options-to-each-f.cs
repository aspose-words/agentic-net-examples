using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure the required folders exist.
        string templatesDir = Path.Combine(Directory.GetCurrentDirectory(), "Templates");
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(templatesDir);
        Directory.CreateDirectory(outputDir);

        // Create sample data model.
        var order = new Order
        {
            CustomerName = "John Doe",
            Items =
            {
                new Item { Index = 1, Name = "Apple" },
                new Item { Index = 2, Name = "Banana" },
                new Item { Index = 3, Name = "Cherry" }
            }
        };

        // Create a few template files programmatically (only once).
        for (int i = 1; i <= 3; i++)
        {
            string templatePath = Path.Combine(templatesDir, $"Template{i}.docx");
            if (!File.Exists(templatePath))
            {
                var doc = new Document();
                var builder = new DocumentBuilder(doc);

                // Header with customer name.
                builder.Writeln($"Report {i}");
                builder.Writeln("Customer: <<[order.CustomerName]>>");
                builder.Writeln();

                // Table header.
                builder.Writeln("<<foreach [item in order.Items]>>");
                var table = builder.StartTable();
                builder.InsertCell();
                builder.Writeln("Index");
                builder.InsertCell();
                builder.Writeln("Product");
                builder.EndRow();

                // Data row.
                builder.InsertCell();
                builder.Writeln("<<[item.Index]>>");
                builder.InsertCell();
                builder.Writeln("<<[item.Name]>>");
                builder.EndRow();
                builder.EndTable();
                builder.Writeln("<</foreach>>");

                doc.Save(templatePath);
            }
        }

        // Process each template file with identical reporting options.
        foreach (string templateFile in Directory.GetFiles(templatesDir, "*.docx"))
        {
            var doc = new Document(templateFile);

            // Configure the reporting engine.
            var engine = new ReportingEngine
            {
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report using the same data model for every template.
            engine.BuildReport(doc, order, "order");

            // Save the generated report.
            string outputPath = Path.Combine(outputDir,
                Path.GetFileNameWithoutExtension(templateFile) + "_Report.docx");
            doc.Save(outputPath);
        }
    }
}

// Data model classes used by the LINQ Reporting engine.
public class Order
{
    public string CustomerName { get; set; } = "";
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = "";
}
