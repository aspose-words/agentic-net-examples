using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare working folders.
        string workDir = Directory.GetCurrentDirectory();
        string imagesDir = Path.Combine(workDir, "Images");
        string outputDir = Path.Combine(workDir, "Output");
        Directory.CreateDirectory(imagesDir);
        Directory.CreateDirectory(outputDir);

        // Create three identical PNG images from a tiny red dot base64 string.
        const string base64Png =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAukB9WcXK7cAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        string[] imagePaths = new[]
        {
            Path.Combine(imagesDir, "image1.png"),
            Path.Combine(imagesDir, "image2.png"),
            Path.Combine(imagesDir, "image3.png")
        };
        foreach (string path in imagePaths)
            File.WriteAllBytes(path, pngBytes);

        // -----------------------------------------------------------------
        // Build a template document that contains a foreach loop and an image tag inside a textbox.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Begin foreach over Items collection.
        builder.Writeln("<<foreach [item in Items]>>");

        // Insert a textbox that will host the image.
        var textBox = builder.InsertShape(Aspose.Words.Drawing.ShapeType.TextBox, 200, 120);
        // Move the cursor inside the textbox.
        builder.MoveTo(textBox.FirstParagraph);
        // Image tag – the expression returns a file path.
        builder.Write("<<image [item.ImagePath] -fitSize>>");
        // Move the cursor back to the main story after the textbox.
        builder.MoveTo(template.FirstSection.Body.LastParagraph);

        // End foreach.
        builder.Writeln("<</foreach>>");

        // -----------------------------------------------------------------
        // Generate a separate report for each image.
        // -----------------------------------------------------------------
        for (int i = 0; i < imagePaths.Length; i++)
        {
            // Prepare data model containing a single item.
            var model = new ReportModel
            {
                Items = new List<ReportItem>
                {
                    new ReportItem { ImagePath = imagePaths[i] }
                }
            };

            // Clone the template so that each report starts from a fresh copy.
            Document report = (Document)template.Clone(true);

            // Build the report using LINQ Reporting Engine.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // default options
            engine.BuildReport(report, model, "model");

            // Save the generated document.
            string outFile = Path.Combine(outputDir, $"Report_{i + 1}.docx");
            report.Save(outFile);
        }

        Console.WriteLine("Reports generated in: " + outputDir);
    }
}

// Root wrapper class referenced in the template as <<[model.Items]>>.
public class ReportModel
{
    public List<ReportItem> Items { get; set; } = new();
}

// Simple data class exposing the image path.
public class ReportItem
{
    public string ImagePath { get; set; } = string.Empty;
}
