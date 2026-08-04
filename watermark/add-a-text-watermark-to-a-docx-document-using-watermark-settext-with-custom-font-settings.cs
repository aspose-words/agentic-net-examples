using System;
using System.Drawing;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Configure custom text watermark options.
        TextWatermarkOptions options = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 48,
            Color = Color.Gray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = false
        };

        // Apply the text watermark with the specified options.
        doc.Watermark.SetText("Confidential", options);

        // Save the resulting document.
        string outputFile = "WatermarkedDocument.docx";
        doc.Save(outputFile);
    }
}
