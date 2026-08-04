using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

public class Item
{
    public string Name { get; set; } = "";
    public int Quantity { get; set; }
    public double Price { get; set; }
}

public class ItemReport
{
    public string Name { get; set; } = "";
    public int Quantity { get; set; }
}

public class ReportModel
{
    public List<ItemReport> Items { get; set; } = new();
}

// Marked as partial to avoid conflict with the test harness's partial declaration.
public partial class Program
{
    public static void Main()
    {
        // 1. Prepare sample data.
        List<Item> items = new()
        {
            new Item { Name = "Apple",  Quantity = 10, Price = 0.5 },
            new Item { Name = "Banana", Quantity = 20, Price = 0.3 },
            new Item { Name = "Orange", Quantity = 15, Price = 0.4 }
        };

        // 2. Project only the fields required for the report using LINQ Select.
        List<ItemReport> projected = items
            .Select(i => new ItemReport { Name = i.Name, Quantity = i.Quantity })
            .ToList();

        // 3. Wrap the projected collection in a root model.
        ReportModel model = new() { Items = projected };

        // 4. Create the LINQ Reporting template programmatically.
        const string templatePath = "Template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Write a simple heading.
        builder.Writeln("Item Report");
        builder.Writeln();

        // Begin a foreach loop over the Items collection.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Name: <<[item.Name]>>");
        builder.Writeln("Quantity: <<[item.Quantity]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // 5. Load the template (ensures BuildReport is called after loading).
        Document loadedTemplate = new Document(templatePath);

        // 6. Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None // No special options needed.
        };
        engine.BuildReport(loadedTemplate, model, "model");

        // 7. Save the generated report.
        const string outputPath = "Report.docx";
        loadedTemplate.Save(outputPath);
    }
}
