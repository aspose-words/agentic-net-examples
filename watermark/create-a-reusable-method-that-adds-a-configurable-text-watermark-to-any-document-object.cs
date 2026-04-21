using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    // Adds a configurable text watermark to the supplied Document.
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
        TextWatermarkOptions options = new TextWatermarkOptions
        {
            FontFamily = fontFamily,
            FontSize = fontSize,
            Color = color ?? Color.Black,
            Layout = layout,
            IsSemitrasparent = isSemitransparent
        };

        // Apply the watermark.
        doc.Watermark.SetText(text, options);
    }

    public static void Main()
    {
        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document.");
        builder.Writeln("It will have a text watermark applied.");

        // Apply a configurable watermark.
        AddTextWatermark(
            doc,
            text: "CONFIDENTIAL",
            fontFamily: "Calibri",
            fontSize: 48f,
            color: Color.Red,
            layout: WatermarkLayout.Diagonal,
            isSemitransparent: false);

        // Ensure output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the watermarked document.
        string outputPath = Path.Combine(outputDir, "Watermarked.docx");
        doc.Save(outputPath);
    }
}
