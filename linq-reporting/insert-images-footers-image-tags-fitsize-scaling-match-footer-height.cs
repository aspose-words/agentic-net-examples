using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

class InsertImageIntoFooter
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Obtain a DocumentBuilder to edit the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the primary footer of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

        // Create a tiny PNG image (1x1 pixel) from a Base64 string.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);
        using var imageStream = new MemoryStream(imageBytes);

        // Insert the image from the stream. The method returns the Shape that contains the image.
        Shape imageShape = builder.InsertImage(imageStream);

        // Set the shape to be a floating picture (no text wrapping) so it can be positioned freely.
        imageShape.WrapType = WrapType.None;
        imageShape.BehindText = true;

        // Position the image relative to the footer margin.
        imageShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Margin;
        imageShape.RelativeVerticalPosition = RelativeVerticalPosition.Margin;

        // Align the image to the left side of the footer.
        imageShape.Left = 0;
        imageShape.Top = 0;

        // Resize the shape to match the height of the footer.
        double footerHeight = builder.PageSetup.BottomMargin;
        imageShape.Height = footerHeight;

        // Set a positive width; using the same value as height preserves aspect ratio for a square image.
        imageShape.Width = footerHeight;

        // Fit the image data to the shape frame so the picture scales to the shape size.
        imageShape.ImageData.FitImageToShape();

        // Determine output path in the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DocumentWithFooterImage.docx");

        // Save the document.
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
