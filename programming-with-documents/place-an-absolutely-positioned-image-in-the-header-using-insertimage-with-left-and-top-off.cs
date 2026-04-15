using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare a tiny PNG image (1x1 pixel, blue) as a byte array.
        // This avoids the need for System.Drawing which may not be available.
        // The image is encoded in base64 for easy inclusion.
        const string base64Png =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ" +
            "b6VYAAAAASUVORK5CYII="; // 1x1 blue pixel

        // Determine the path where the image will be saved.
        string imagePath = Path.Combine(Directory.GetCurrentDirectory(), "sample.png");

        // Write the PNG bytes to the file system.
        File.WriteAllBytes(imagePath, Convert.FromBase64String(base64Png));

        // Create a new blank Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the primary header of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Insert the image into the header. The image is inserted as an inline shape initially.
        Shape shape = builder.InsertImage(imagePath);

        // Convert the shape to a floating image.
        shape.WrapType = WrapType.None;
        shape.BehindText = true;

        // Position the image absolutely relative to the page.
        shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        shape.Left = 50; // points from the left edge of the page
        shape.Top = 20;  // points from the top edge of the page

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "HeaderImage.docx");
        doc.Save(outputPath);
    }
}
