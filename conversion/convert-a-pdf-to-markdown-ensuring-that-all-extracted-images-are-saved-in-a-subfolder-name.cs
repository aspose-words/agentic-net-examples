using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names and folders.
        string inputPdfPath = Path.Combine(Directory.GetCurrentDirectory(), "sample.pdf");
        string outputMarkdownPath = Path.Combine(Directory.GetCurrentDirectory(), "sample.md");
        string assetsFolder = Path.Combine(Directory.GetCurrentDirectory(), "assets");

        // -----------------------------------------------------------------
        // 1. Create a sample PDF document with an image.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample PDF that will be converted to Markdown.");
        // Insert a sample image (use any image file that exists in the project directory).
        // For demonstration, create a simple placeholder image file if it does not exist.
        string placeholderImagePath = Path.Combine(Directory.GetCurrentDirectory(), "placeholder.png");
        if (!File.Exists(placeholderImagePath))
        {
            // Create a 1x1 pixel PNG using Aspose.Drawing (no System.Drawing usage).
            using (var bitmap = new Aspose.Drawing.Bitmap(1, 1))
            {
                bitmap.SetPixel(0, 0, Aspose.Drawing.Color.Red);
                bitmap.Save(placeholderImagePath, Aspose.Drawing.Imaging.ImageFormat.Png);
            }
        }
        builder.InsertImage(placeholderImagePath);
        // Save the document as PDF.
        sourceDoc.Save(inputPdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(inputPdfPath))
            throw new InvalidOperationException("Failed to create the input PDF file.");

        // -----------------------------------------------------------------
        // 2. Load the PDF and convert it to Markdown, saving images to "assets".
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(inputPdfPath);

        // Configure Markdown save options.
        MarkdownSaveOptions mdOptions = new MarkdownSaveOptions
        {
            // Ensure images are saved to the "assets" subfolder.
            ImagesFolder = assetsFolder,
            ImagesFolderAlias = "assets",
            SaveFormat = SaveFormat.Markdown
        };

        // Save as Markdown.
        pdfDoc.Save(outputMarkdownPath, mdOptions);

        // -----------------------------------------------------------------
        // 3. Validation.
        // -----------------------------------------------------------------
        if (!File.Exists(outputMarkdownPath))
            throw new InvalidOperationException("Markdown output file was not created.");

        if (!Directory.Exists(assetsFolder))
            throw new InvalidOperationException("Assets folder was not created.");

        // At least one image file should be present in the assets folder.
        string[] imageFiles = Directory.GetFiles(assetsFolder);
        if (imageFiles.Length == 0)
            throw new InvalidOperationException("No images were extracted to the assets folder.");

        // Example completed successfully.
        Console.WriteLine("PDF successfully converted to Markdown.");
        Console.WriteLine($"Markdown file: {outputMarkdownPath}");
        Console.WriteLine($"Extracted images folder: {assetsFolder}");
    }
}
