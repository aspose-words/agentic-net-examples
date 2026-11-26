using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for sample PNG images.
        string imagesFolder = "SampleImages";
        Directory.CreateDirectory(imagesFolder);

        // Create a few PNG images using Aspose.Drawing.
        int imageCount = 3;
        string[] imagePaths = new string[imageCount];
        for (int i = 0; i < imageCount; i++)
        {
            string filePath = Path.Combine(imagesFolder, $"Image{i + 1}.png");
            using (Bitmap bitmap = new Bitmap(600, 800))
            {
                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    // Fill background with a distinct color.
                    Color background = i % 2 == 0 ? Color.LightBlue : Color.LightGreen;
                    graphics.Clear(background);
                }

                // Save the bitmap as PNG.
                bitmap.Save(filePath, ImageFormat.Png);
            }

            imagePaths[i] = filePath;
        }

        // Create a new Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert each PNG on a separate page.
        for (int i = 0; i < imagePaths.Length; i++)
        {
            if (i > 0)
            {
                // Add a page break before the next image.
                builder.InsertBreak(BreakType.PageBreak);
            }

            // Insert the image; let Aspose.Words scale it to fit the page.
            builder.InsertImage(imagePaths[i]);
        }

        // Save the document as a single PDF.
        string outputPdf = "CombinedImages.pdf";
        doc.Save(outputPdf, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(outputPdf))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }
    }
}
