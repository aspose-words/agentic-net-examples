using System;
using System.Drawing;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Configure custom text watermark options.
        TextWatermarkOptions options = new TextWatermarkOptions
        {
            FontFamily = "Arial",          // Font family for the watermark text.
            FontSize = 48,                 // Font size in points.
            Color = Color.Blue,            // Text color.
            Layout = WatermarkLayout.Diagonal, // Diagonal layout.
            IsSemitrasparent = false       // Opaque watermark.
        };

        // Add the text watermark with the specified options.
        doc.Watermark.SetText("Confidential", options);

        // Save the resulting document to the local file system.
        string outputPath = "WatermarkedDocument.docx";
        doc.Save(outputPath);
    }
}
