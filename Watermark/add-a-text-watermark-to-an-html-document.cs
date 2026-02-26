using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source HTML document.
        string inputPath = "input.html";

        // Path where the watermarked HTML will be saved.
        string outputPath = "output.html";

        // Load the HTML document.
        Document doc = new Document(inputPath);

        // Configure watermark appearance.
        TextWatermarkOptions watermarkOptions = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 36,
            Color = Color.Gray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = false
        };

        // Apply the text watermark to the document.
        doc.Watermark.SetText("Confidential", watermarkOptions);

        // Save the document back to HTML format.
        doc.Save(outputPath, SaveFormat.Html);
    }
}
