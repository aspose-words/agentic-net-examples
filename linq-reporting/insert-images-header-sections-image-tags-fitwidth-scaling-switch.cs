using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertHeaderImage
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Create a DocumentBuilder which will be used to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the primary header of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Insert the image into the header. The method returns the Shape that contains the image.
        // Replace the path with the actual location of your image file.
        string imagePath = @"C:\Images\logo.png";
        Shape imageShape = builder.InsertImage(imagePath);

        // Make the image behave like a watermark (no text wrapping, placed behind text).
        imageShape.WrapType = WrapType.None;
        imageShape.BehindText = true;

        // Position the image relative to the page.
        imageShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        imageShape.RelativeVerticalPosition = RelativeVerticalPosition.Page;

        // Align the image to the left edge of the page.
        imageShape.Left = 0;

        // Scale the image to fit the page width while preserving its aspect ratio.
        double targetWidth = builder.PageSetup.PageWidth;
        double scaleFactor = targetWidth / imageShape.Width;
        imageShape.Width = targetWidth;
        imageShape.Height = imageShape.Height * scaleFactor;

        // Save the document.
        doc.Save(@"C:\Output\HeaderImage.docx");
    }
}
