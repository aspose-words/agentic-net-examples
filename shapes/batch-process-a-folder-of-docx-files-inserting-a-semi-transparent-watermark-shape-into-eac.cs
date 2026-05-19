using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class BatchWatermarkProcessor
{
    public static void Main()
    {
        // Define input and output directories relative to the current working directory.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputDocs");

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Process each DOCX file in the input folder.
        foreach (string inputPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(inputPath);

            // Insert a semi‑transparent rectangle shape as a watermark into the primary header.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

            // Create a floating rectangle shape.
            Shape watermark = builder.InsertShape(ShapeType.Rectangle, 500, 500); // size in points

            // Position the shape behind the text and center it on the page.
            watermark.WrapType = WrapType.None;
            watermark.BehindText = true;
            watermark.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            watermark.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            watermark.HorizontalAlignment = HorizontalAlignment.Center;
            watermark.VerticalAlignment = VerticalAlignment.Center;

            // Set a light gray fill color with 50 % transparency.
            watermark.FillColor = Color.LightGray;
            watermark.Fill.Transparency = 0.5; // 0 = opaque, 1 = fully transparent

            // Remove the shape outline.
            watermark.StrokeColor = Color.Transparent;

            // Save the modified document to the output folder, preserving the original file name.
            string outputPath = Path.Combine(outputFolder, Path.GetFileName(inputPath));
            doc.Save(outputPath, SaveFormat.Docx);

            // Validate that the file was saved successfully.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to save the watermarked document: {outputPath}");
        }
    }
}
