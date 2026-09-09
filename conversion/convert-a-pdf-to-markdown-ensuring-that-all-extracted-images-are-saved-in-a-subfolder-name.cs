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
        // Define paths for the sample files and folders.
        string workingDir = Directory.GetCurrentDirectory();
        string pdfPath = Path.Combine(workingDir, "sample.pdf");
        string markdownPath = Path.Combine(workingDir, "sample.md");
        string assetsFolder = Path.Combine(workingDir, "assets");
        string imagePath = Path.Combine(workingDir, "sample.png");

        // Ensure the assets folder exists.
        Directory.CreateDirectory(assetsFolder);

        // Create a simple PNG image using Aspose.Drawing.
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Blue);
            }
            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // Create a Word document, add text and the image, then save it as PDF.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample PDF containing an image.");
        builder.InsertImage(imagePath);
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Configure Markdown save options to store images in the "assets" subfolder.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            ImagesFolder = assetsFolder,
            SaveFormat = SaveFormat.Markdown
        };

        // Convert the PDF to Markdown.
        pdfDoc.Save(markdownPath, saveOptions);

        // Validation: ensure the Markdown file was created.
        if (!File.Exists(markdownPath))
            throw new InvalidOperationException("The Markdown output file was not created.");

        // Validation: ensure at least one image was saved to the assets folder.
        if (!Directory.Exists(assetsFolder))
            throw new InvalidOperationException("The assets folder does not exist.");

        string[] extractedImages = Directory.GetFiles(assetsFolder);
        if (extractedImages.Length == 0)
            throw new InvalidOperationException("No images were extracted to the assets folder.");

        // Optional: clean up temporary files (comment out if you want to inspect the results).
        // File.Delete(imagePath);
        // File.Delete(pdfPath);
        // File.Delete(markdownPath);
        // foreach (string file in extractedImages) File.Delete(file);
        // Directory.Delete(assetsFolder);
    }
}
