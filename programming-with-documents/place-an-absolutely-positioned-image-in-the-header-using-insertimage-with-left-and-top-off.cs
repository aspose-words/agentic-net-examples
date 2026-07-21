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

        // A tiny 1x1 pixel PNG (transparent) encoded in Base64.
        // This avoids the need for System.Drawing dependencies.
        const string base64Png = 
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ" +
            "Z6XcAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);

        // Insert the image as a floating shape using the byte[] overload.
        // Width and height are set to -1 to keep the original image dimensions.
        Shape shape = builder.InsertImage(imageBytes, -1, -1);

        // Configure the shape to be absolutely positioned in the header.
        shape.WrapType = WrapType.None;                     // No text wrapping.
        shape.BehindText = true;                            // Place behind the text.
        shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page; // Position relative to page.
        shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;     // Position relative to page.
        shape.Left = 50;   // Left offset in points.
        shape.Top = 30;    // Top offset in points.

        // Save the document.
        doc.Save("HeaderImage.docx");
    }
}
