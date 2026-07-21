using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Define input and output folders relative to the current directory.
        string inputFolder = Path.Combine(Environment.CurrentDirectory, "InputDocs");
        string outputFolder = Path.Combine(Environment.CurrentDirectory, "OutputDocs");

        // Ensure both directories exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // If the input folder is empty, create a sample DOCX file to process.
        if (Directory.GetFiles(inputFolder, "*.docx").Length == 0)
        {
            Document sample = new Document();
            DocumentBuilder sampleBuilder = new DocumentBuilder(sample);
            sampleBuilder.Writeln("This is a sample document for watermark processing.");
            string samplePath = Path.Combine(inputFolder, "Sample.docx");
            sample.Save(samplePath);
        }

        // Process each DOCX file in the input folder.
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(filePath);
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Move to the primary header so the watermark appears on every page.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

            // Insert a floating rectangle shape that will serve as the watermark.
            Shape watermark = builder.InsertShape(ShapeType.Rectangle, 300, 200);
            watermark.FillColor = Color.LightGray;
            // Set semi‑transparent fill (0 = opaque, 1 = fully transparent).
            watermark.Fill.Transparency = 0.5f;

            // Position the shape behind the text and center it on the page.
            watermark.WrapType = WrapType.None;
            watermark.BehindText = true;
            watermark.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            watermark.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            watermark.HorizontalAlignment = HorizontalAlignment.Center;
            watermark.VerticalAlignment = VerticalAlignment.Center;

            // Save the modified document to the output folder.
            string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
            doc.Save(outputPath);

            // Verify that the file was saved successfully.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to save processed file: {outputPath}");
        }
    }
}
