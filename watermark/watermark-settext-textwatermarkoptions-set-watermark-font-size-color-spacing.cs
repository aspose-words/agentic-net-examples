using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

class WatermarkExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Configure watermark options.
        TextWatermarkOptions options = new TextWatermarkOptions();
        options.FontSize = 48;               // Set the watermark font size.
        options.Color = Color.DarkGray;      // Set the watermark font color.
        // TextWatermarkOptions does not provide a direct property for character spacing.
        // If spacing is required, a custom shape with a TextPath would be needed.

        // Apply the text watermark with the specified options.
        doc.Watermark.SetText("Confidential", options);

        // Save the document.
        doc.Save("WatermarkedDocument.docx");
    }
}
