using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class ImageData
{
    // Image bytes (PNG, JPEG, etc.)
    public byte[] Data { get; set; } = Array.Empty<byte>();
}

public class ReportModel
{
    // Collection of images to be inserted.
    public List<ImageData> Images { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample images (small PNGs encoded in Base64).
        string base64Png = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAQAAAC1+jfqAAAAFklEQVR4AWP8z8Dwn4EIwDiqgAIAAP//AwB6VgVZAAAAAElFTkSuQmCC"; // 16x16 red square

        var model = new ReportModel
        {
            Images = new List<ImageData>
            {
                new ImageData { Data = Convert.FromBase64String(base64Png) },
                new ImageData { Data = Convert.FromBase64String(base64Png) }
            }
        };

        // -----------------------------------------------------------------
        // Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Begin a foreach loop over the Images collection.
        builder.Writeln("<<foreach [img in Images]>>");

        // Create a table to hold each image in its own row.
        Table table = builder.StartTable();
        builder.InsertCell();

        // Insert a textbox that will act as the image container.
        // Width = 200 points, Height = 150 points (container size).
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 150);
        builder.MoveTo(textBox.FirstParagraph);

        // Insert the image tag with -fitHeight switch to scale height while preserving width proportionally.
        builder.Write("<<image [img.Data] -fitHeight>>");

        // Return to the main document flow (the cell's paragraph).
        builder.MoveTo(table.LastRow.LastCell.LastParagraph);

        // Finish the table row and table.
        builder.EndRow();
        builder.EndTable();

        // Close the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report using the model as the root object named "model".
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(outputPath);
    }
}
