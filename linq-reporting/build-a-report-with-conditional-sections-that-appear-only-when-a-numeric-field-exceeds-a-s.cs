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
        // Register code page provider (required for some Aspose.Words features)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data
        ReportModel model = new()
        {
            Threshold = 50,
            Items = new()
            {
                new Item { Name = "Item A", Value = 30 },
                new Item { Name = "Item B", Value = 60 },
                new Item { Name = "Item C", Value = 45 },
                new Item { Name = "Item D", Value = 80 }
            }
        };

        // Create a template document programmatically
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        CreateTemplate(templatePath);

        // Load the template
        Document doc = new(templatePath);

        // Build the report using LINQ Reporting Engine
        ReportingEngine engine = new()
        {
            // Remove empty paragraphs that may appear after processing tags
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };
        engine.BuildReport(doc, model, "model");

        // Save the generated report
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ReportOutput.docx");
        doc.Save(outputPath);
    }

    private static void CreateTemplate(string filePath)
    {
        Document template = new();
        DocumentBuilder builder = new(template);

        // Header
        builder.Writeln("Report Summary");
        builder.Writeln("Threshold: <<[model.Threshold]>>");

        // Iterate over items
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Item: <<[item.Name]>> Value: <<[item.Value]>>");

        // Conditional section: appears only when the item's value exceeds the threshold
        builder.Writeln("<<if [item.Value > model.Threshold]>> - Exceeds threshold<</if>>");

        // End of foreach
        builder.Writeln("<</foreach>>");

        // Save the template
        template.Save(filePath);
    }
}

// Root data model
public class ReportModel
{
    public int Threshold { get; set; } = 0;
    public List<Item> Items { get; set; } = new();
}

// Item model used in the collection
public class Item
{
    public string Name { get; set; } = string.Empty;
    public int Value { get; set; }
}
