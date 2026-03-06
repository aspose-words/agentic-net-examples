using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Configure text watermark options.
        TextWatermarkOptions watermarkOptions = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 36,
            Color = Color.Gray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = false
        };

        // Add a text watermark to every page of the document.
        doc.Watermark.SetText("Confidential", watermarkOptions);

        // Save the document in DOCX format.
        doc.Save("WatermarkedDocument.docx");
    }
}
