using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class WatermarkDemo
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Configure text watermark options.
        TextWatermarkOptions textOptions = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 36,
            Color = Color.Black,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = false
        };

        // Add a text watermark to the document.
        doc.Watermark.SetText("Confidential", textOptions);

        // Save the document that now contains the watermark.
        doc.Save("WatermarkedText.docx");

        // Load the previously saved document.
        Document loadedDoc = new Document("WatermarkedText.docx");

        // If the document has a text watermark, remove it.
        if (loadedDoc.Watermark.Type == WatermarkType.Text)
        {
            loadedDoc.Watermark.Remove();
        }

        // Save the document after the watermark has been removed.
        loadedDoc.Save("WatermarkRemoved.docx");
    }
}
