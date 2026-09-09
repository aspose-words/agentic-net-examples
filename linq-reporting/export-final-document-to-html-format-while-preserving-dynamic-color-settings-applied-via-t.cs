using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Folder for generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a Word template with LINQ Reporting tags.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(outputDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin a foreach loop over the Items collection.
        builder.Writeln("<<foreach [item in Items]>>");

        // Apply a dynamic text color using the <<textColor>> tag.
        // The color expression will be taken from item.Color.
        builder.Writeln("<<textColor [item.Color]>>");
        builder.Writeln("<<[item.Text]>>");
        builder.Writeln("<</textColor>>");

        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and prepare the data model.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // Sample data model.
        ReportModel model = new()
        {
            Items = new()
            {
                new Item { Text = "First line - red",   Color = "Red" },
                new Item { Text = "Second line - green", Color = "Green" },
                new Item { Text = "Third line - blue",  Color = "Blue" }
            }
        };

        // -----------------------------------------------------------------
        // 3. Build the report using the LINQ Reporting engine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // default options
        engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // 4. Save the final document as HTML, preserving the colors.
        // -----------------------------------------------------------------
        HtmlSaveOptions htmlOptions = new()
        {
            // Export colors as they appear in the document.
            ExportFontResources = true,
            ExportImagesAsBase64 = true,
            // Optional: specify a folder for external resources (not needed for base64).
            ImagesFolder = Path.Combine(outputDir, "Images")
        };

        string htmlPath = Path.Combine(outputDir, "Report.html");
        doc.Save(htmlPath, htmlOptions);

        Console.WriteLine($"Report generated: {htmlPath}");
    }
}

// ---------------------------------------------------------------------
// Data model classes used by the LINQ Reporting engine.
// ---------------------------------------------------------------------
public class ReportModel
{
    // Collection referenced by the template's foreach tag.
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    // Text to display.
    public string Text { get; set; } = string.Empty;

    // Color name or HTML color code used by the <<textColor>> tag.
    public string Color { get; set; } = string.Empty;
}
