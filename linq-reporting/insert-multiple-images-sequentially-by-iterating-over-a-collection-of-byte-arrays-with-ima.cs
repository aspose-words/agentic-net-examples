using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a template document with LINQ Reporting tags.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Begin foreach over the Images collection.
        builder.Writeln("<<foreach [img in Images]>>");

        // Table that will hold one image per row.
        builder.StartTable();
        builder.InsertCell();

        // Insert a textbox shape inside the cell to host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 150, 150);
        // Move the cursor into the textbox's first paragraph.
        builder.MoveTo(textBox.FirstParagraph);
        // Image tag that references the byte‑array property.
        builder.Writeln("<<image [img.Data]>>");

        // Return the cursor to the main story (outside the textbox) for table continuation.
        builder.MoveTo(template.LastSection.Body.LastParagraph);

        // End the table row and the table.
        builder.EndRow();
        builder.EndTable();

        // Close the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template locally.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Prepare the data model with sample images.
        ReportModel model = new ReportModel
        {
            Images = new List<ImageItem>
            {
                new ImageItem { Data = Convert.FromBase64String(SampleBase64Png1) },
                new ImageItem { Data = Convert.FromBase64String(SampleBase64Png2) }
            }
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(reportDoc, model, "model");

        // Save the final document.
        reportDoc.Save("Report.docx");
    }

    // Simple data model required by the template.
    public class ReportModel
    {
        public List<ImageItem> Images { get; set; } = new();
    }

    public class ImageItem
    {
        public byte[] Data { get; set; } = Array.Empty<byte>();
    }

    // Two tiny 1×1 PNG images encoded in Base64.
    private const string SampleBase64Png1 =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ" +
        "Zc6cAAAAASUVORK5CYII="; // Red pixel

    private const string SampleBase64Png2 =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8AAAgEB" +
        "A6V6WQAAAABJRU5ErkJggg=="; // Blue pixel
}
