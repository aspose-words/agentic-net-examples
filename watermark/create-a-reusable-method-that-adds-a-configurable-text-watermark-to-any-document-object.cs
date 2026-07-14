using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    // Adds a configurable text watermark to the provided Document.
    public static void AddTextWatermark(
        Document doc,
        string text,
        string fontFamily = "Arial",
        float fontSize = 36f,
        Color? color = null,
        WatermarkLayout layout = WatermarkLayout.Diagonal,
        bool isSemitransparent = false)
    {
        if (doc == null) throw new ArgumentNullException(nameof(doc));
        if (string.IsNullOrWhiteSpace(text)) throw new ArgumentException("Watermark text cannot be null or whitespace.", nameof(text));

        var options = new TextWatermarkOptions
        {
            FontFamily = fontFamily,
            FontSize = fontSize,
            Color = color ?? Color.Black,
            Layout = layout,
            IsSemitrasparent = isSemitransparent
        };

        doc.Watermark.SetText(text, options);
    }

    public static void Main()
    {
        // Create a blank document and add some sample content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document.");
        builder.Writeln("The watermark should appear behind this text.");

        // Apply a configurable text watermark.
        AddTextWatermark(
            doc,
            text: "CONFIDENTIAL",
            fontFamily: "Calibri",
            fontSize: 48f,
            color: Color.Red,
            layout: WatermarkLayout.Diagonal,
            isSemitransparent: false);

        // Save the result.
        const string outputPath = "WatermarkedDocument.docx";
        doc.Save(outputPath);
    }
}
