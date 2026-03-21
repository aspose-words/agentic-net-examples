using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Base64ImageInsertion
{
    static void Main()
    {
        // Base64-encoded PNG image (1x1 pixel, transparent)
        string base64Image = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X9ZcAAAAASUVORK5CYII=";

        // Convert the Base64 string to a byte array.
        byte[] imageBytes = Convert.FromBase64String(base64Image);

        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image using the byte array. This creates an inline shape with the original image dimensions.
        builder.InsertImage(imageBytes);

        // Insert a page break to demonstrate multiple images.
        builder.InsertBreak(BreakType.PageBreak);

        // Insert the same image with custom dimensions (250x144 pixels).
        double width = ConvertUtil.PixelToPoint(250);
        double height = ConvertUtil.PixelToPoint(144);
        builder.InsertImage(imageBytes, width, height);

        // Save the resulting document.
        doc.Save("Base64Images.docx");
    }
}
