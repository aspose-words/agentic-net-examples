using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Define output folder and file.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "ReportWithWatermark.docx");

        // Create a new blank document.
        Document doc = new Document();

        // Optionally add some sample content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Automated Report");
        builder.Writeln("Generated on: " + DateTime.Now);

        // Configure text watermark options.
        TextWatermarkOptions watermarkOptions = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 48,
            Color = Color.Red,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = false
        };

        // Apply the confidential text watermark.
        doc.Watermark.SetText("CONFIDENTIAL", watermarkOptions);

        // Save the document.
        doc.Save(outputPath);

        // Simple verification that the file was created.
        Console.WriteLine(File.Exists(outputPath)
            ? $"Document saved successfully to: {outputPath}"
            : "Failed to save the document.");
    }
}
