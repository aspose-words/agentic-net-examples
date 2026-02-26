using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Load the DOTX template
        Document doc = new Document("Template.dotx");

        // Optional: define watermark appearance
        TextWatermarkOptions options = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 36,
            Color = Color.Gray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = true
        };

        // Add a text watermark to the document
        doc.Watermark.SetText("Confidential", options);

        // Save the modified document, preserving the DOTX format
        doc.Save("Template_Watermarked.dotx");
    }
}
