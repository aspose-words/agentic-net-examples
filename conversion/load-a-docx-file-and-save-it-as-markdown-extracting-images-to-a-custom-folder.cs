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
        // Paths for temporary files and output
        const string docxPath = "input.docx";
        const string markdownPath = "output.md";
        const string imagePath = "sample.png";
        const string imagesFolder = "output_images";

        // Ensure a clean environment
        if (File.Exists(docxPath)) File.Delete(docxPath);
        if (File.Exists(markdownPath)) File.Delete(markdownPath);
        if (File.Exists(imagePath)) File.Delete(imagePath);
        if (Directory.Exists(imagesFolder)) Directory.Delete(imagesFolder, true);

        // Create a simple PNG image using Aspose.Drawing
        using (Bitmap bitmap = new Bitmap(100, 100, PixelFormat.Format32bppArgb))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Blue);
            }
            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // Create a sample DOCX document containing text and the image
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample document with an image:");
        builder.InsertImage(imagePath);
        sourceDoc.Save(docxPath, SaveFormat.Docx);

        // Load the DOCX document
        Document doc = new Document(docxPath);

        // Prepare Markdown save options with a custom images folder
        MarkdownSaveOptions markdownOptions = new MarkdownSaveOptions
        {
            ImagesFolder = imagesFolder,
            ImagesFolderAlias = imagesFolder // Use the same folder name in the markdown links
        };

        // Ensure the images folder exists
        Directory.CreateDirectory(imagesFolder);

        // Save as Markdown, extracting images to the custom folder
        doc.Save(markdownPath, markdownOptions);

        // Validation: check that the markdown file was created
        if (!File.Exists(markdownPath))
            throw new InvalidOperationException("The Markdown file was not created.");

        // Validation: check that at least one image was extracted
        string[] extractedImages = Directory.GetFiles(imagesFolder);
        if (extractedImages.Length == 0)
            throw new InvalidOperationException("No images were extracted to the custom folder.");

        // Optional: clean up temporary files (comment out if inspection is needed)
        // File.Delete(docxPath);
        // File.Delete(imagePath);
        // Directory.Delete(imagesFolder, true);
    }
}
