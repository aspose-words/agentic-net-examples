using Aspose.Words;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document (or load an existing RTF with new Document("input.rtf"))
        Document doc = new Document();

        // Configure watermark appearance
        TextWatermarkOptions options = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 36,
            Color = Color.Gray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = false
        };

        // Add the text watermark to the document
        doc.Watermark.SetText("Confidential", options);

        // Save the document as RTF
        doc.Save("Output.rtf");
    }
}
