using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertHeaderImage
{
    static void Main()
    {
        // Create a temporary image file (a simple red square) so the example runs without external resources.
        string tempImagePath = Path.Combine(Path.GetTempPath(), "tempLogo.png");
        // Base64-encoded 2x2 red PNG.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAADUlEQVQImWNgYGD4DwABBAEAj6cK5QAAAABJRU5ErkJggg==";
        File.WriteAllBytes(tempImagePath, Convert.FromBase64String(base64Png));

        // Create a new empty document.
        Document doc = new Document();

        // Create a DocumentBuilder which will be used to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the primary header of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Insert the image into the header.
        Shape imageShape = builder.InsertImage(tempImagePath);

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
        double originalWidth = imageShape.Width;
        double originalHeight = imageShape.Height;
        double scaleFactor = targetWidth / originalWidth;
        imageShape.Width = targetWidth;
        imageShape.Height = originalHeight * scaleFactor;

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "HeaderImage.docx");

        // Save the document.
        doc.Save(outputPath);

        // Clean up the temporary image file.
        File.Delete(tempImagePath);
    }
}
