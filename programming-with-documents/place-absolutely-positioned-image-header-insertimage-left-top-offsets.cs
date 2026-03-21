using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Prepare a temporary PNG image (1x1 pixel transparent).
        string tempImagePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".png");
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X0ZcAAAAASUVORK5CYII=";
        File.WriteAllBytes(tempImagePath, Convert.FromBase64String(base64Png));

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the builder's cursor to the primary header of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Insert a floating image into the header with absolute positioning.
        Shape shape = builder.InsertImage(
            tempImagePath,
            RelativeHorizontalPosition.Page,
            50, // left offset (points)
            RelativeVerticalPosition.Page,
            30, // top offset (points)
            -1, // width (auto)
            -1, // height (auto)
            WrapType.None);

        // Ensure the image appears behind any overlapping text.
        shape.BehindText = true;

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the resulting document.
        doc.Save(Path.Combine(outputDir, "HeaderImage.docx"));

        // Clean up the temporary image file.
        File.Delete(tempImagePath);
    }
}
