using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // URL of the online video to embed.
        string videoUrl = "https://example.com/video.mp4";

        // Insert the OLE object as a linked package.
        // progId "Package" is used for generic file containers.
        // isLink = true creates a link to the online resource.
        // The last parameter (imageFileName) is set to null because we do not provide a placeholder image.
        Shape oleShape = builder.InsertOleObject("Package", videoUrl, true, false, null);

        // Set precise width and height (in points). 1 inch = 72 points.
        // Example: 5 inches wide and 3 inches high.
        oleShape.Width = ConvertUtil.InchToPoint(5);
        oleShape.Height = ConvertUtil.InchToPoint(3);

        // Save the document to disk.
        doc.Save("Output.docx");
    }
}
