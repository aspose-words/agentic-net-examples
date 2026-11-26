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
        // Prepare temporary paths
        string tempFolder = Path.Combine(Path.GetTempPath(), "AsposeSample");
        Directory.CreateDirectory(tempFolder);
        string imageFilePath = Path.Combine(tempFolder, "sample_image.png");
        string pdfPath = Path.Combine(tempFolder, "sample.pdf");
        string markdownPath = Path.Combine(tempFolder, "output.md");
        string imagesFolder = Path.Combine(tempFolder, "images");

        // Create a simple bitmap image using Aspose.Drawing
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Blue);
            }
            bitmap.Save(imageFilePath, ImageFormat.Png);
        }

        // Build a sample document and insert the image
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample PDF content with an image:");
        builder.InsertImage(imageFilePath);
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF document
        Document pdfDoc = new Document(pdfPath);

        // Configure Markdown save options to extract images to a folder with relative paths
        MarkdownSaveOptions mdOptions = new MarkdownSaveOptions
        {
            ImagesFolder = imagesFolder,
            ImagesFolderAlias = "images"
        };

        // Ensure the images folder exists
        Directory.CreateDirectory(imagesFolder);

        // Save as Markdown
        pdfDoc.Save(markdownPath, mdOptions);

        // Validation
        if (!File.Exists(markdownPath))
            throw new InvalidOperationException("Markdown file was not created.");

        if (!Directory.Exists(imagesFolder))
            throw new InvalidOperationException("Images folder was not created.");

        string[] extractedImages = Directory.GetFiles(imagesFolder);
        if (extractedImages.Length == 0)
            throw new InvalidOperationException("No images were extracted to the images folder.");

        // Clean up temporary image file used for insertion (optional)
        if (File.Exists(imageFilePath))
            File.Delete(imageFilePath);
    }
}
