using System.Drawing;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Configure the appearance of the text watermark.
        TextWatermarkOptions watermarkOptions = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 36,
            Color = Color.Black,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = false
        };

        // Add the text watermark to the document.
        doc.Watermark.SetText("Aspose Watermark", watermarkOptions);

        // Save the document as a DOCX file.
        doc.Save("Document.TextWatermark.docx");
    }
}
