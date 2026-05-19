using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a template document with LINQ Reporting tags.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Order Report");
        builder.Writeln("<<foreach [item in Items]>>");
        // This line will become an empty paragraph if item.Name is empty.
        builder.Writeln("Item: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Step 2: Load the template document.
        Document doc = new Document(templatePath);

        // Step 3: Prepare sample data with some empty values.
        ReportModel model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Apple" },
                new Item { Name = "" },          // This will generate an empty paragraph.
                new Item { Name = "Banana" },
                new Item { Name = null },        // This will also generate an empty paragraph.
                new Item { Name = "Cherry" }
            }
        };

        // Step 4: Build the report with the option to remove empty paragraphs.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;
        engine.BuildReport(doc, model, "model");

        // Step 5: Save the final document.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}

// Public data model classes required by the template.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    // Initialize to empty string to avoid nullable warnings.
    public string Name { get; set; } = string.Empty;
}
