using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    // Reusable method that adds a configurable text watermark to any Document.
    public static void AddTextWatermark(
        Document doc,
        string text,
        string fontFamily = "Arial",
        float fontSize = 36f,
        Color? color = null,
        WatermarkLayout layout = WatermarkLayout.Diagonal,
        bool isSemitransparent = false)
    {
        // Prepare watermark options.
        var options = new TextWatermarkOptions
        {
            FontFamily = fontFamily,
            FontSize = fontSize,
            Color = color ?? Color.Gray,
            Layout = layout,
            IsSemitrasparent = isSemitransparent
        };

        // Apply the watermark.
        doc.Watermark.SetText(text, options);
    }

    public static void Main()
    {
        // Ensure output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a blank document and add some sample content.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document.");
        builder.Writeln("The text watermark will appear behind this text.");

        // Add a configurable text watermark.
        AddTextWatermark(
            doc,
            text: "CONFIDENTIAL",
            fontFamily: "Calibri",
            fontSize: 48f,
            color: Color.Red,
            layout: WatermarkLayout.Diagonal,
            isSemitransparent: true);

        // Save the watermarked document.
        string outputPath = Path.Combine(outputDir, "Watermarked.docx");
        doc.Save(outputPath);

        // Simple validation that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Watermarked document saved to: {outputPath}");
        }
        else
        {
            Console.WriteLine("Failed to save the watermarked document.");
        }
    }
}
