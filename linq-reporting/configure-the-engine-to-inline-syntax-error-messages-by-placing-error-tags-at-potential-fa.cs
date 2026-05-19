using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data model.
        var order = new Order
        {
            CustomerName = "John Doe",
            Items = new List<Item>
            {
                new Item { Name = "Apple", Quantity = 3 },
                new Item { Name = "Banana", Quantity = 5 }
            }
        };

        // Create a template document programmatically.
        var templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Load the template.
        var doc = new Document(templatePath);

        // Configure the reporting engine to inline error messages.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report. The boolean indicates whether parsing succeeded.
        bool success = engine.BuildReport(doc, order, "order");

        // Save the resulting document.
        var outputPath = "Report_Output.docx";
        doc.Save(outputPath);

        // Simple console output to indicate completion.
        Console.WriteLine($"Report generation completed. Success flag: {success}");
        Console.WriteLine($"Output saved to: {Path.GetFullPath(outputPath)}");
    }

    private static void CreateTemplate(string filePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Normal data insertion.
        builder.Writeln("Customer: <<[order.CustomerName]>>");

        // Loop over items.
        builder.Writeln("<<foreach [item in order.Items]>>");
        builder.Writeln("Item: <<[item.Name]>>  Qty: <<[item.Quantity]>>");
        builder.Writeln("<</foreach>>");

        // Intentional syntax error: malformed expression inside an if tag.
        builder.Writeln("<<if [order.Items.Count >]>>");
        builder.Writeln("This line will be skipped due to syntax error.");
        builder.Writeln("<</if>>");

        // Reference to a missing property (will also trigger an error message).
        builder.Writeln("Missing property: <<[order.MissingProperty]>>");

        // Save the template.
        doc.Save(filePath);
    }
}

// Data model classes.
public class Order
{
    public string CustomerName { get; set; } = "";
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = "";
    public int Quantity { get; set; }
}
