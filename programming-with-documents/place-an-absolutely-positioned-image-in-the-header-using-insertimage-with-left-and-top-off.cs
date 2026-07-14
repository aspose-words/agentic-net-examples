using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace HeaderImageExample
{
    public class Program
    {
        public static void Main()
        {
            // Prepare a tiny PNG image (1x1 pixel, transparent) from a Base64 string.
            // This avoids the need for System.Drawing dependencies.
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAukB9YV6cVYAAAAASUVORK5CYII=";
            byte[] imageBytes = Convert.FromBase64String(base64Png);
            string imagePath = Path.Combine(Directory.GetCurrentDirectory(), "tempImage.png");
            File.WriteAllBytes(imagePath, imageBytes);

            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Move the cursor to the primary header.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

            // Insert the image into the header.
            Shape shape = builder.InsertImage(imagePath);

            // Configure the shape as a floating image with absolute offsets.
            shape.WrapType = WrapType.None;                     // No text wrapping.
            shape.BehindText = true;                            // Appear behind the text.
            shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            shape.Left = 100;                                   // Left offset in points.
            shape.Top = 50;                                     // Top offset in points.

            // Save the document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "HeaderImage.docx");
            doc.Save(outputPath);

            // Clean up the temporary image file.
            if (File.Exists(imagePath))
                File.Delete(imagePath);
        }
    }
}
