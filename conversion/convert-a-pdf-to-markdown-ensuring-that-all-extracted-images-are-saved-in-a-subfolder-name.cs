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
        // Define working directories and file names.
        string workDir = Directory.GetCurrentDirectory();
        string assetsFolder = Path.Combine(workDir, "assets");
        string imagePath = Path.Combine(workDir, "sample.png");
        string pdfPath = Path.Combine(workDir, "sample.pdf");
        string markdownPath = Path.Combine(workDir, "output.md");

        // Ensure the assets folder exists (Aspose.Words will also create it if missing).
        if (!Directory.Exists(assetsFolder))
            Directory.CreateDirectory(assetsFolder);

        // -----------------------------------------------------------------
        // 1. Create a simple image using Aspose.Drawing (no System.Drawing usage).
        // -----------------------------------------------------------------
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            // Use the static FromImage method to obtain a Graphics object.
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Blue);
            }
            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // -----------------------------------------------------------------
        // 2. Create a sample PDF that contains the image.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample PDF with an embedded image.");
        builder.InsertImage(imagePath);
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // -----------------------------------------------------------------
        // 3. Load the PDF and convert it to Markdown, saving images to "assets".
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        MarkdownSaveOptions mdOptions = new MarkdownSaveOptions
        {
            // Physical folder where images will be written.
            ImagesFolder = assetsFolder,
            // URI part used in the Markdown file to reference images.
            ImagesFolderAlias = "assets",
            // Explicitly set the format to Markdown.
            SaveFormat = SaveFormat.Markdown
        };

        pdfDoc.Save(markdownPath, mdOptions);

        // -----------------------------------------------------------------
        // 4. Validation: ensure Markdown file and extracted images exist.
        // -----------------------------------------------------------------
        if (!File.Exists(markdownPath))
            throw new InvalidOperationException("Markdown file was not created.");

        string[] savedImages = Directory.GetFiles(assetsFolder);
        if (savedImages.Length == 0)
            throw new InvalidOperationException("No images were extracted to the assets folder.");

        // Output paths for verification (no interactive prompts).
        Console.WriteLine("Conversion completed successfully.");
        Console.WriteLine("Markdown file: " + markdownPath);
        Console.WriteLine("Extracted images folder: " + assetsFolder);
    }
}
