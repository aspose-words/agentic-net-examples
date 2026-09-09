using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices; // Required for the assembly attribute
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables; // Provides the Table class used in the example

// This attribute makes internal members of this assembly visible to the Aspose.Words.Reporting assembly,
// which is required when merging the reporting engine with another assembly for deployment.
[assembly: InternalsVisibleTo("Aspose.Words.Reporting")]

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Order = new Order
            {
                CustomerName = "John Doe",
                Items = new List<Item>
                {
                    new Item { Index = 1, Name = "Apple" },
                    new Item { Index = 2, Name = "Banana" },
                    new Item { Index = 3, Name = "Cherry" }
                }
            }
        };

        // Create a template document programmatically.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Insert LINQ Reporting tags.
        builder.Writeln("Customer: <<[model.Order.CustomerName]>>");
        builder.Writeln("<<foreach [item in model.Order.Items]>>");

        // Build a simple table inside the foreach block.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Index");
        builder.InsertCell();
        builder.Writeln("Product");
        builder.EndRow();

        builder.InsertCell();
        builder.Writeln("<<[item.Index]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.EndRow();
        builder.EndTable();

        builder.Writeln("<</foreach>>");

        // Save the template (optional, shown for clarity).
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template and build the report.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        reportDoc.Save("Report.docx");
    }
}

// Root wrapper class for the report.
public class ReportModel
{
    public Order Order { get; set; } = new();
}

// Order class containing customer info and a collection of items.
public class Order
{
    public string CustomerName { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

// Simple item class.
public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
}
