using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the primary header of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // A 1x1 pixel PNG image (transparent) encoded in Base64.
        // This avoids the need for System.Drawing dependencies.
        string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2V8AAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);

        // Insert the image as a floating shape with absolute left/top offsets.
        Shape shape = builder.InsertImage(
            imageBytes,
            RelativeHorizontalPosition.Page, 100,   // 100 points from the left edge of the page
            RelativeVerticalPosition.Page, 50,      // 50 points from the top edge of the page
            -1, -1,                                 // Use the image's original width and height
            WrapType.None);

        // Ensure the image is placed behind any text.
        shape.BehindText = true;
        shape.WrapType = WrapType.None;

        // Save the resulting document.
        doc.Save("HeaderImage.docx");
    }
}
