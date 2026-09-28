using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class Item
{
    public string Name { get; set; } = "";
    public string? Description { get; set; }
}

public class Order
{
    public string CustomerName { get; set; } = "";
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features).
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        const string templatePath = "Template.docx";
        const string outputPath = "Report.docx";

        // -------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Title paragraph.
        builder.Writeln("Order Report for <<[order.CustomerName]>>");
        builder.Writeln();

        // Begin foreach over Items.
        builder.Writeln("<<foreach [item in Items]>>");

        // Paragraph with item name.
        builder.Writeln("Item: <<[item.Name]>>");

        // Conditional paragraph that may be empty if Description is null.
        builder.Writeln("<<if [item.Description != null]>>Description: <<[item.Description]>> <</if>>");

        // End foreach.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template and prepare sample data.
        // -------------------------------------------------
        var doc = new Document(templatePath);

        var order = new Order
        {
            CustomerName = "Acme Corp",
            Items = new()
            {
                new Item { Name = "Widget", Description = "A useful widget." },
                new Item { Name = "Gadget", Description = null }, // No description – will produce an empty paragraph.
                new Item { Name = "Doohickey", Description = "An interesting doohickey." }
            }
        };

        // -------------------------------------------------
        // 3. Build the report with the RemoveEmptyParagraphs option.
        // -------------------------------------------------
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;
        engine.BuildReport(doc, order, "order");

        // -------------------------------------------------
        // 4. Save the final report.
        // -------------------------------------------------
        doc.Save(outputPath);
    }
}
