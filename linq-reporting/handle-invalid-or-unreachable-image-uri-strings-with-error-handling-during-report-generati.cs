using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // -----------------------------------------------------------------
        // 1. Create a tiny PNG image (1x1 pixel) to be used as a valid image.
        // -----------------------------------------------------------------
        string validImagePath = Path.Combine(workDir, "validImage.png");
        // Base64 for a 1x1 transparent PNG.
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=");
        File.WriteAllBytes(validImagePath, pngBytes);

        // -----------------------------------------------------------------
        // 2. Build the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(workDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Report with images (invalid URIs will show error messages):");
        // Begin a foreach block over Items collection.
        builder.Writeln("<<foreach [item in Items]>>");

        // Insert a textbox that will host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
        builder.MoveTo(textBox.FirstParagraph);
        // Image tag – the expression returns a string (file path or URI).
        builder.Write("<<image [item.ImageUri] -fitSize>>");

        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare the data model with one valid and one invalid image URI.
        // -----------------------------------------------------------------
        var data = new ReportData();
        data.Items.Add(new ReportItem { ImageUri = validImagePath }); // valid local file.
        data.Items.Add(new ReportItem { ImageUri = "http://nonexistent.example.com/missing.png" }); // invalid URI.

        // -----------------------------------------------------------------
        // 4. Load the template and build the report with inline error messages.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine
        {
            // InlineErrorMessages makes the engine insert error text instead of throwing.
            Options = ReportBuildOptions.InlineErrorMessages
        };

        // BuildReport returns true if parsing succeeded (errors are inlined).
        bool success = engine.BuildReport(reportDoc, data, "data");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        string outputPath = Path.Combine(workDir, "ReportOutput.docx");
        reportDoc.Save(outputPath);

        // Indicate completion (no interactive prompts).
        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}.");
        Console.WriteLine($"Output saved to: {outputPath}");
    }
}

// ---------------------------------------------------------------------
// Data model definitions.
// ---------------------------------------------------------------------
public class ReportItem
{
    // ImageUri can be a file path or a web URL.
    public string ImageUri { get; set; } = "";
}

public class ReportData
{
    // Collection that the template iterates over.
    public List<ReportItem> Items { get; set; } = new();
}
