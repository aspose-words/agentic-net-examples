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
        string inputDocxPath = Path.Combine(workingDir, "sample.docx");
        string outputMarkdownPath = Path.Combine(workingDir, "output.md");
        string imagesFolderPath = Path.Combine(workingDir, "Images");
        string sampleImagePath = Path.Combine(workingDir, "sample.png");

        // Ensure the images folder exists.
        Directory.CreateDirectory(imagesFolderPath);

        // -----------------------------------------------------------------
        // Create a simple PNG image using Aspose.Drawing (no System.Drawing).
        // -----------------------------------------------------------------
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill the bitmap with a solid color.
                graphics.Clear(Color.Blue);
            }

            // Save the bitmap to a file that will be inserted into the DOCX.
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // ---------------------------------------------------------------
        // Create a sample DOCX document containing some text and the image.
        // ---------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document with an image.");
        builder.InsertImage(sampleImagePath);
        doc.Save(inputDocxPath, SaveFormat.Docx);

        // ---------------------------------------------------------------
        // Load the DOCX file that we just created.
        // ---------------------------------------------------------------
        Document loadedDoc = new Document(inputDocxPath);

        // ---------------------------------------------------------------
        // Configure Markdown save options to extract images to a custom folder.
        // ---------------------------------------------------------------
        MarkdownSaveOptions markdownOptions = new MarkdownSaveOptions
        {
            ImagesFolder = imagesFolderPath,
            SaveFormat = SaveFormat.Markdown
        };

        // Save the document as Markdown. Images will be written to ImagesFolder.
        loadedDoc.Save(outputMarkdownPath, markdownOptions);

        // ---------------------------------------------------------------
        // Validation: ensure the Markdown file and extracted images exist.
        // ---------------------------------------------------------------
        if (!File.Exists(outputMarkdownPath))
            throw new InvalidOperationException("The Markdown output file was not created.");

        string[] extractedImages = Directory.GetFiles(imagesFolderPath);
        if (extractedImages.Length == 0)
            throw new InvalidOperationException("No images were extracted to the specified folder.");

        // Optional: indicate successful completion.
        Console.WriteLine("Conversion completed successfully.");
        Console.WriteLine($"Markdown file: {outputMarkdownPath}");
        Console.WriteLine($"Extracted images folder: {imagesFolderPath}");
    }
}
