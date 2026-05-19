using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    // Simple data model used by the template.
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

    public static void Main()
    {
        // Create a blank document and a builder to insert LINQ Reporting tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Template content.
        builder.Writeln("Customer: <<[model.CustomerName]>>");
        builder.Writeln("<<foreach [item in model.Items]>>");
        builder.Writeln("- <<[item.Index]>>: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        // Prepare sample data.
        Order order = new Order
        {
            CustomerName = "Acme Corp",
            Items = new List<Item>
            {
                new Item { Index = 1, Name = "Widget" },
                new Item { Index = 2, Name = "Gadget" }
            }
        };

        // Configure the reporting engine to inline error messages.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report and capture the success flag.
        bool success = engine.BuildReport(doc, order, "model");

        // Output the result flag.
        Console.WriteLine($"Report build success: {success}");

        // Save the generated document.
        doc.Save("ReportOutput.docx");
    }
}
