using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Define watermark options: font size, color, and layout.
        TextWatermarkOptions options = new TextWatermarkOptions
        {
            FontSize = 48,               // Set the font size of the watermark.
            Color = Color.Blue,          // Set the watermark color.
            FontFamily = "Arial",        // Optional: set a specific font family.
            Layout = WatermarkLayout.Diagonal, // Layout of the watermark.
            IsSemitrasparent = false    // Make the watermark fully opaque.
        };

        // Apply the text watermark with the specified options.
        doc.Watermark.SetText("CONFIDENTIAL", options);

        // Save the document to a file in the current directory.
        string outputPath = "WatermarkedDocument.docx";
        doc.Save(outputPath);
    }
}
