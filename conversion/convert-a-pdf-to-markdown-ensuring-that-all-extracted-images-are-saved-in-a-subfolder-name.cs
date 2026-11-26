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
        // Define paths.
        string assetsFolder = "assets";
        string imagePath = "sample.png";
        string pdfPath = "input.pdf";
        string markdownPath = "output.md";

        // Ensure the assets folder exists.
        Directory.CreateDirectory(assetsFolder);

        // Create a simple image using Aspose.Drawing.
        Bitmap bitmap = new Bitmap(100, 100);
        try
        {
            Graphics graphics = Graphics.FromImage(bitmap);
            try
            {
                graphics.Clear(Color.Blue);
            }
            finally
            {
                graphics.Dispose();
            }

            // Save the image to a file.
            bitmap.Save(imagePath, ImageFormat.Png);
        }
        finally
        {
            bitmap.Dispose();
        }

        // Create a sample document, insert text and the image, then save as PDF.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample PDF document containing an image.");
        builder.InsertImage(imagePath);
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Configure Markdown save options to store images in the assets folder.
        MarkdownSaveOptions mdOptions = new MarkdownSaveOptions
        {
            ImagesFolder = assetsFolder,
            ImagesFolderAlias = assetsFolder
        };

        // Save the PDF as Markdown.
        pdfDoc.Save(markdownPath, mdOptions);

        // Validation: ensure Markdown file exists.
        if (!File.Exists(markdownPath))
        {
            throw new InvalidOperationException("The Markdown output file was not created.");
        }

        // Validation: ensure at least one image was extracted to the assets folder.
        string[] extractedImages = Directory.GetFiles(assetsFolder);
        if (extractedImages.Length == 0)
        {
            throw new InvalidOperationException("No images were extracted to the assets folder.");
        }

        // Clean up temporary files (optional).
        // File.Delete(imagePath);
        // File.Delete(pdfPath);
        // Note: The assets folder and its contents are retained for inspection.
    }
}
