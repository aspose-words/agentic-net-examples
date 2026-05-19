using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // 1. Create the LINQ Reporting template.
        string templatePath = Path.Combine(workDir, "Template.docx");
        CreateTemplate(templatePath);

        // 2. Load the template for reporting.
        Document template = new Document(templatePath);

        // 3. Prepare sample data.
        ReportModel model = new()
        {
            Items = new()
            {
                new Item { Name = "Apple", Index = 1 },
                new Item { Name = "Banana", Index = 2 },
                new Item { Name = "Cherry", Index = 3 }
            }
        };

        // 4. Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // 5. Save the generated report.
        string outputPath = Path.Combine(workDir, "Report.docx");
        template.Save(outputPath);
    }

    // Creates a Word document containing LINQ Reporting tags.
    private static void CreateTemplate(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Title.
        builder.Writeln("Items with first‑character styling:");
        builder.Writeln();

        // Begin foreach loop over Items.
        builder.Writeln("<<foreach [item in Items]>>");

        // Apply a red color to the first character of the Name,
        // then output the remaining characters normally.
        builder.Writeln(
            "<<textColor [\"Red\"]>><<[item.Name.Substring(0,1)]>><</textColor>><<[item.Name.Substring(1)]>>");

        // New line for each item.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(filePath);
    }
}

// Root data model for the report.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Simple item class used in the foreach loop.
public class Item
{
    public string Name { get; set; } = string.Empty;
    public int Index { get; set; }
}
