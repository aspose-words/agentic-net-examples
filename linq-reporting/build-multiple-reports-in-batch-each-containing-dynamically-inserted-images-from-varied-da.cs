using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare working folders.
        string baseDir = Directory.GetCurrentDirectory();
        string templateDir = Path.Combine(baseDir, "Template");
        string imagesDir = Path.Combine(baseDir, "Images");
        string outputDir = Path.Combine(baseDir, "Output");
        Directory.CreateDirectory(templateDir);
        Directory.CreateDirectory(imagesDir);
        Directory.CreateDirectory(outputDir);

        // Create sample PNG images.
        CreateSampleImage(Path.Combine(imagesDir, "image1.png"));
        CreateSampleImage(Path.Combine(imagesDir, "image2.png"));
        CreateSampleImage(Path.Combine(imagesDir, "image3.png"));

        // Build the LINQ Reporting template and save it to disk.
        string templatePath = Path.Combine(templateDir, "ReportTemplate.docx");
        BuildTemplate(templatePath);

        // Prepare batch data – a list of report models.
        var batchModels = new List<ReportModel>
        {
            new()
            {
                Items =
                {
                    new() { Title = "First Item", ImagePath = Path.Combine(imagesDir, "image1.png") },
                    new() { Title = "Second Item", ImagePath = Path.Combine(imagesDir, "image2.png") }
                }
            },
            new()
            {
                Items =
                {
                    new() { Title = "Third Item", ImagePath = Path.Combine(imagesDir, "image3.png") },
                    new() { Title = "Fourth Item", ImagePath = Path.Combine(imagesDir, "image1.png") }
                }
            }
        };

        // Generate each report in the batch.
        for (int i = 0; i < batchModels.Count; i++)
        {
            // Load a fresh copy of the template for each report.
            var doc = new Document(templatePath);

            // Build the report using LINQ Reporting Engine.
            var engine = new ReportingEngine();
            engine.BuildReport(doc, batchModels[i], "model");

            // Save the generated report.
            string outputPath = Path.Combine(outputDir, $"Report_{i + 1}.docx");
            doc.Save(outputPath);
        }
    }

    // Creates a minimal 1x1 PNG image from a Base64 string.
    private static void CreateSampleImage(string filePath)
    {
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(filePath, imageBytes);
    }

    // Builds the Word template with LINQ Reporting tags.
    private static void BuildTemplate(string templatePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Batch Report");
        builder.Writeln("<<foreach [item in Items]>>");

        // Table header.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Title");
        builder.InsertCell();
        builder.Writeln("Image");
        builder.EndRow();

        // Data row.
        builder.InsertCell();
        builder.Writeln("<<[item.Title]>>");
        builder.InsertCell();

        // Insert a textbox to host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [item.ImagePath] -fitSize>>");

        builder.EndRow();
        builder.EndTable();

        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);
    }
}

// Data model for a single report item.
public class ReportItem
{
    public string Title { get; set; } = "";
    public string ImagePath { get; set; } = "";
}

// Root data model for a report.
public class ReportModel
{
    public List<ReportItem> Items { get; set; } = new();
}
