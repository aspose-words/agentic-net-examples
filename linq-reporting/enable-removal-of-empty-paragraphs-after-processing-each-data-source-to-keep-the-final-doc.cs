using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        string templatePath = "Template.docx";
        string reportPath = "Report.docx";

        // 1. Create the LINQ Reporting template programmatically.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a foreach block that iterates over Items.
        builder.Writeln("<<foreach [item in Items]>>");
        // Each iteration writes the item's Name on its own paragraph.
        builder.Writeln("<<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // 2. Load the template for report generation.
        Document loadedTemplate = new Document(templatePath);

        // 3. Prepare the data source.
        ReportModel model = new()
        {
            Items = new()
            {
                new Item { Name = "Alice" },
                new Item { Name = string.Empty }, // This will produce an empty paragraph.
                new Item { Name = "Bob" }
            }
        };

        // 4. Configure the ReportingEngine to remove empty paragraphs.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

        // 5. Build the report.
        engine.BuildReport(loadedTemplate, model, "model");

        // 6. Save the final document.
        loadedTemplate.Save(reportPath);
    }
}

// Root data model for the report.
public class ReportModel
{
    // Collection of items to be iterated in the template.
    public List<Item> Items { get; set; } = new();
}

// Simple item class referenced by the template.
public class Item
{
    // Name may be empty; initialize to avoid nullable warnings.
    public string Name { get; set; } = string.Empty;
}
