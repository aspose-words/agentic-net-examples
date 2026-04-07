using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Define output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        string templatePath = Path.Combine(outputDir, "Template.docx");
        string resultPath = Path.Combine(outputDir, "Result.docx");

        // -----------------------------------------------------------------
        // 1. Create a template document using the default LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Use the standard << >> delimiters required by the LINQ Reporting engine.
        builder.Writeln("<<[model.Title]>>");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln(" - <<[item.Name]>>: <<[item.Quantity]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Prepare sample data model.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel
        {
            Title = "Product Inventory",
            Items = new List<Item>
            {
                new Item { Name = "Apples", Quantity = 120 },
                new Item { Name = "Bananas", Quantity = 85 },
                new Item { Name = "Oranges", Quantity = 60 }
            }
        };

        // -----------------------------------------------------------------
        // 3. Load the template and build the report.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // Build the report using the model as the root object named "model".
        engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        doc.Save(resultPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes (public with initialized properties to avoid warnings).
// ---------------------------------------------------------------------
public class ReportModel
{
    public string Title { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public int Quantity { get; set; }
}
