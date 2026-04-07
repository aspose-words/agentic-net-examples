using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for Aspose.Words on .NET Core).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        Order order = new()
        {
            CustomerName = "John Doe",
            Items =
            {
                new Item { Index = 1, Name = "Apple" },
                new Item { Index = 2, Name = "Banana" },
                new Item { Index = 3, Name = "Cherry" }
            }
        };

        // Create a template document programmatically.
        string templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Load the template.
        Document doc = new(templatePath);

        // Build the report using contextual member access inside the foreach band.
        ReportingEngine engine = new();
        engine.BuildReport(doc, order, "order");

        // Save the generated report.
        string reportPath = "Report.docx";
        doc.Save(reportPath);
    }

    // Creates a Word document containing LINQ Reporting tags.
    private static void CreateTemplate(string filePath)
    {
        Document doc = new();
        DocumentBuilder builder = new(doc);

        // Header with a root object reference.
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln();

        // Data band (foreach) that uses contextual member access.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Item <<[Index]>>: <<[Name]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(filePath);
    }
}

// Root data model.
public class Order
{
    public string CustomerName { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

// Item model used inside the data band.
public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
}
