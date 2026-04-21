using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class BatchWatermarkProcessor
{
    public static void Main()
    {
        // Prepare input and output folders.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputDocs");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample DOCX files to demonstrate the batch process.
        CreateSampleDocuments(inputDir);

        // Process each DOCX file: load, add watermark, and save.
        foreach (string filePath in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document doc = new Document(filePath);
            InsertSemiTransparentWatermark(doc);
            string outputPath = Path.Combine(outputDir, Path.GetFileName(filePath));
            doc.Save(outputPath);

            // Validate that the output file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to save watermarked document: {outputPath}");
        }
    }

    private static void CreateSampleDocuments(string folder)
    {
        for (int i = 1; i <= 3; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"This is sample document #{i}.");
            string fileName = Path.Combine(folder, $"Sample{i}.docx");
            doc.Save(fileName);
        }
    }

    private static void InsertSemiTransparentWatermark(Document doc)
    {
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Determine page dimensions.
        double pageWidth = builder.PageSetup.PageWidth;
        double pageHeight = builder.PageSetup.PageHeight;

        // Insert a floating rectangle that covers the whole page.
        Shape watermark = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.Page, 0,
            RelativeVerticalPosition.Page, 0,
            pageWidth,
            pageHeight,
            WrapType.None);

        // Place the shape behind the text.
        watermark.BehindText = true;

        // Apply a semi‑transparent fill color.
        watermark.FillColor = Color.FromArgb(80, Color.LightGray); // 80/255 opacity.

        // Remove the outline.
        watermark.Stroke.Color = Color.Transparent;
    }
}
