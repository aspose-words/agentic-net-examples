using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        const string pdfPath = "sample.pdf";
        const string jpgPath = "sample.jpg";
        const string pngPath = "sample.png";

        // -----------------------------------------------------------------
        // 1. Create a sample Word document and save it as PDF.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample document.");
        builder.InsertImage(ImageDirPlaceholder()); // Placeholder for an image if needed.
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify PDF creation.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // -----------------------------------------------------------------
        // 2. Load the PDF document.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // -----------------------------------------------------------------
        // 3. Save the PDF as JPEG.
        // -----------------------------------------------------------------
        pdfDoc.Save(jpgPath, SaveFormat.Jpeg);
        if (!File.Exists(jpgPath))
            throw new InvalidOperationException("JPEG image was not created.");

        // -----------------------------------------------------------------
        // 4. Save the PDF as PNG.
        // -----------------------------------------------------------------
        pdfDoc.Save(pngPath, SaveFormat.Png);
        if (!File.Exists(pngPath))
            throw new InvalidOperationException("PNG image was not created.");

        // All conversions succeeded.
        Console.WriteLine("Conversion completed successfully.");
    }

    // Helper method to provide a valid image path for the sample document.
    // In a real scenario, replace this with an actual image file path.
    private static string ImageDirPlaceholder()
    {
        // Create a temporary 1x1 pixel PNG image if it does not exist.
        const string tempImage = "placeholder.png";
        if (!File.Exists(tempImage))
        {
            using (MemoryStream ms = new MemoryStream())
            {
                // Create a minimal PNG using Aspose.Words image saving (single white pixel).
                Document imgDoc = new Document();
                DocumentBuilder imgBuilder = new DocumentBuilder(imgDoc);
                imgBuilder.Writeln(" ");
                ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
                {
                    ImageColorMode = ImageColorMode.None,
                    Resolution = 72,
                    ImageSize = new System.Drawing.Size(1, 1) // System.Drawing is not used; size is ignored for placeholder.
                };
                imgDoc.Save(ms, options);
                File.WriteAllBytes(tempImage, ms.ToArray());
            }
        }
        return tempImage;
    }
}
