using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create a sample image using Aspose.Drawing and save it to a file.
        string imagePath = "sample.png";
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.LightBlue);
                using (Pen pen = new Pen(Color.DarkBlue, 2))
                {
                    graphics.DrawRectangle(pen, new Rectangle(10, 10, 80, 80));
                }
            }
            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // Create a sample Word document, insert the image, and save it as PDF.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample PDF document with an image.");
        builder.InsertImage(imagePath);
        string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Ensure the assets folder exists for extracted images.
        string assetsFolder = "assets";
        Directory.CreateDirectory(assetsFolder);

        // Load the PDF and convert it to Markdown, extracting images to the assets folder.
        Document pdfDoc = new Document(pdfPath);
        MarkdownSaveOptions mdOptions = new MarkdownSaveOptions
        {
            ImagesFolder = assetsFolder,
            ImagesFolderAlias = assetsFolder
        };
        string markdownPath = "output.md";
        pdfDoc.Save(markdownPath, mdOptions);

        // Validate that the Markdown file was created.
        if (!File.Exists(markdownPath))
            throw new InvalidOperationException("Markdown file was not created.");

        // Validate that at least one image was extracted to the assets folder.
        string[] extractedImages = Directory.GetFiles(assetsFolder);
        if (extractedImages.Length == 0)
            throw new InvalidOperationException("No images were extracted to the assets folder.");
    }
}
