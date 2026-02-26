using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load the existing DOCX document that contains OfficeMath objects.
        Document doc = new Document("Input.docx");

        // Configure the appearance of the text watermark.
        TextWatermarkOptions watermarkOptions = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 36,
            Color = Color.LightGray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = true
        };

        // Apply the watermark to the entire document (it will appear behind all content,
        // including any OfficeMath equations).
        doc.Watermark.SetText("CONFIDENTIAL", watermarkOptions);

        // Save the watermarked document.
        doc.Save("Output.docx");
    }
}
