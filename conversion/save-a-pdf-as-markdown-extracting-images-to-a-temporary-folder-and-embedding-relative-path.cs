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
        // Define temporary working directories.
        string baseTempDir = Path.Combine(Path.GetTempPath(), "AsposeConversionExample");
        Directory.CreateDirectory(baseTempDir);

        string imagesFolder = Path.Combine(baseTempDir, "images");
        Directory.CreateDirectory(imagesFolder);

        string pdfPath = Path.Combine(baseTempDir, "sample.pdf");
        string markdownPath = Path.Combine(baseTempDir, "sample.md");
        string sampleImagePath = Path.Combine(baseTempDir, "sample.png");

        // Create a simple image using Aspose.Drawing.
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Blue);
            }
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // Create a sample document containing text and the image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample document with an image:");
        builder.InsertImage(sampleImagePath);

        // Save the document as PDF.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Configure Markdown save options to extract images to a folder and use relative paths.
        MarkdownSaveOptions mdOptions = new MarkdownSaveOptions
        {
            ImagesFolder = imagesFolder,          // Physical folder where images will be saved.
            ImagesFolderAlias = "images",         // Relative path used in the Markdown file.
            SaveFormat = SaveFormat.Markdown
        };

        // Save the PDF as Markdown.
        pdfDoc.Save(markdownPath, mdOptions);

        // Validate that the Markdown file was created.
        if (!File.Exists(markdownPath))
            throw new InvalidOperationException("The Markdown output file was not created.");

        // Validate that at least one image was extracted.
        if (!Directory.Exists(imagesFolder) || Directory.GetFiles(imagesFolder).Length == 0)
            throw new InvalidOperationException("No images were extracted to the images folder.");

        // Verify that the Markdown content contains the relative image path.
        string markdownContent = File.ReadAllText(markdownPath);
        if (!markdownContent.Contains("images/"))
            throw new InvalidOperationException("The Markdown file does not contain relative image paths.");

        // Example completed successfully.
    }
}
