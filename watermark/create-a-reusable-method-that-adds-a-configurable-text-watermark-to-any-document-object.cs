using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a blank document and add sample text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document.");

        // Apply a configurable text watermark.
        AddTextWatermark(doc, "CONFIDENTIAL", "Arial", 48, Color.Red, WatermarkLayout.Diagonal, false);

        // Save the watermarked document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Watermarked.docx");
        doc.Save(outputPath);
    }

    // Reusable method to add a text watermark with custom settings.
    public static void AddTextWatermark(Document doc, string text, string fontFamily, float fontSize, Color color, WatermarkLayout layout, bool isSemitransparent)
    {
        TextWatermarkOptions options = new TextWatermarkOptions
        {
            FontFamily = fontFamily,
            FontSize = fontSize,
            Color = color,
            Layout = layout,
            IsSemitrasparent = isSemitransparent
        };
        doc.Watermark.SetText(text, options);
    }
}
