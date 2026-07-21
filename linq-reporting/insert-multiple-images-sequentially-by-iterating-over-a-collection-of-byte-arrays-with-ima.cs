using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class ImageItem
{
    // Byte array containing the image data.
    public byte[] ImageBytes { get; set; } = Array.Empty<byte>();
}

public class ReportModel
{
    // Collection of images to be inserted.
    public List<ImageItem> Images { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Simple heading.
        builder.Writeln("Image Report");
        builder.Writeln();

        // Begin a foreach loop over the Images collection.
        builder.Writeln("<<foreach [img in Images]>>");

        // Create a table row that will be repeated for each image.
        Table table = builder.StartTable();
        builder.InsertCell();

        // Insert a textbox that will host the image.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        // Move the cursor inside the textbox.
        builder.MoveTo(textBox.FirstParagraph);
        // Image tag referencing the byte[] property; fit the image to the textbox size.
        builder.Write("<<image [img.ImageBytes] -fitSize>>");

        // Close the table row and table.
        builder.EndRow();
        builder.EndTable();

        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Prepare sample image data (two 1x1 PNG images encoded in Base64).
        // -----------------------------------------------------------------
        // A 1x1 pixel transparent PNG.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAukB9Y6XK6cAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);

        var model = new ReportModel();
        model.Images.Add(new ImageItem { ImageBytes = pngBytes });
        model.Images.Add(new ImageItem { ImageBytes = pngBytes }); // Add a second image for demonstration.

        // -----------------------------------------------------------------
        // 3. Load the template, build the report, and save the result.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
