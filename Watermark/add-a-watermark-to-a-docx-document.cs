using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class WatermarkExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Configure text watermark appearance.
        TextWatermarkOptions watermarkOptions = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 36,
            Color = Color.Gray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = true
        };

        // Add the text watermark to the document.
        doc.Watermark.SetText("Confidential", watermarkOptions);

        // Save the document as a DOCX file.
        doc.Save("Watermarked.docx", new OoxmlSaveOptions(SaveFormat.Docx));
    }
}
