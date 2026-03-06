using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load the DOT template.
        Document doc = new Document("Template.dot");

        // Configure text watermark options.
        TextWatermarkOptions options = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 36,
            Color = Color.Gray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = true
        };

        // Apply the watermark to the document (appears behind all content, including table cells).
        doc.Watermark.SetText("CONFIDENTIAL", options);

        // Save the result.
        doc.Save("Result.docx");
    }
}
