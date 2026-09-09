using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class Program
{
    // Entry point of the console application.
    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Step 1: Create a DOCX template with LINQ Reporting tags.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        CreateTemplate(templatePath);

        // Step 2: Prepare sample data model.
        ReportModel model = new ReportModel
        {
            CustomerName = "John Doe",
            Items = new List<Item>
            {
                new Item { Index = 1, Name = "Apple" },
                new Item { Index = 2, Name = "Banana" },
                new Item { Index = 3, Name = "Cherry" }
            }
        };

        // Step 3: Load the template and build the report.
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        // The root object name used in the template is "model".
        engine.BuildReport(reportDoc, model, "model");

        // Step 4: Save the populated document as PDF.
        string pdfPath = Path.Combine(outputDir, "Report.pdf");
        reportDoc.Save(pdfPath, SaveFormat.Pdf);
    }

    // Creates a simple DOCX file containing LINQ Reporting tags.
    private static void CreateTemplate(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a title.
        builder.Writeln("Customer Order Report");
        builder.Writeln();

        // Insert a placeholder for the customer's name.
        builder.Writeln("Customer: <<[model.CustomerName]>>");
        builder.Writeln();

        // Begin a foreach loop over the collection of items.
        builder.Writeln("<<foreach [item in model.Items]>>");
        // Each iteration writes the item's index and name.
        builder.Writeln("Item <<[item.Index]>>: <<[item.Name]>>");
        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to the specified path.
        doc.Save(filePath);
    }
}

// Root data model referenced by the template (named "model").
public class ReportModel
{
    public string CustomerName { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

// Simple item class used in the collection.
public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
}
