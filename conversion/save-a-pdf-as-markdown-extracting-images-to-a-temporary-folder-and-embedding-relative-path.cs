using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary working directory.
        string baseTemp = Path.Combine(Path.GetTempPath(), "AsposeExample_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(baseTemp);

        // Create a sample PNG image using Aspose.Drawing.
        string imagePath = Path.Combine(baseTemp, "sample.png");
        using (Bitmap bitmap = new Bitmap(200, 200, PixelFormat.Format32bppArgb))
        {
            // Obtain a Graphics object for drawing on the bitmap.
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background with a light color.
                graphics.Clear(Color.FromArgb(255, 173, 216, 230));

                // Draw a simple ellipse.
                using (Pen pen = new Pen(Color.FromArgb(255, 0, 120, 215), 5))
                {
                    graphics.DrawEllipse(pen, new RectangleF(20, 20, 160, 160));
                }
            }

            // Save the bitmap as PNG.
            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // Create a sample PDF document that contains the image.
        string pdfPath = Path.Combine(baseTemp, "input.pdf");
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample PDF content with an embedded image:");
        builder.InsertImage(imagePath);
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The input PDF was not created.");

        // Load the PDF for conversion.
        Document pdfDoc = new Document(pdfPath);

        // Prepare a folder for extracted images.
        string imagesFolder = Path.Combine(baseTemp, "images");
        Directory.CreateDirectory(imagesFolder);

        // Configure Markdown save options to extract images to the folder.
        MarkdownSaveOptions mdOptions = new MarkdownSaveOptions
        {
            ExportImagesAsBase64 = false,
            ImagesFolder = imagesFolder,
            ImagesFolderAlias = "images"
        };

        // Save the PDF as Markdown.
        string markdownPath = Path.Combine(baseTemp, "output.md");
        pdfDoc.Save(markdownPath, mdOptions);

        // Validate that the Markdown file was created.
        if (!File.Exists(markdownPath))
            throw new InvalidOperationException("The Markdown output file was not created.");

        // Validate that at least one image was extracted.
        string[] extractedImages = Directory.GetFiles(imagesFolder);
        if (extractedImages.Length == 0)
            throw new InvalidOperationException("No images were extracted to the image folder.");

        // Example completed successfully.
    }
}
