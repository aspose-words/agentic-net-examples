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
        // Define file and folder names.
        const string inputDocx = "input.docx";
        const string outputMd = "output.md";
        const string imagesFolder = "Images";

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX document with some text and an image.
        // -----------------------------------------------------------------
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Hello World!");

        // Create a simple PNG image using Aspose.Drawing (no System.Drawing usage).
        const string tempImagePath = "sample.png";
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Blue);
            }
            bitmap.Save(tempImagePath, ImageFormat.Png);
        }

        // Insert the generated image into the document.
        builder.InsertImage(tempImagePath);

        // Save the sample document as DOCX (bootstrap step).
        sampleDoc.Save(inputDocx, SaveFormat.Docx);

        // Clean up the temporary image file used only for building the sample.
        if (File.Exists(tempImagePath))
            File.Delete(tempImagePath);

        // -----------------------------------------------------------------
        // 2. Load the DOCX file.
        // -----------------------------------------------------------------
        Document doc = new Document(inputDocx);

        // -----------------------------------------------------------------
        // 3. Configure Markdown save options to extract images to a folder.
        // -----------------------------------------------------------------
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            ImagesFolder = imagesFolder
        };

        // Ensure the images folder exists.
        Directory.CreateDirectory(imagesFolder);

        // -----------------------------------------------------------------
        // 4. Save the document as Markdown.
        // -----------------------------------------------------------------
        doc.Save(outputMd, saveOptions);

        // -----------------------------------------------------------------
        // 5. Validation.
        // -----------------------------------------------------------------
        if (!File.Exists(outputMd))
            throw new InvalidOperationException("The Markdown output file was not created.");

        if (!Directory.Exists(imagesFolder))
            throw new InvalidOperationException("The images folder was not created.");

        string[] extractedImages = Directory.GetFiles(imagesFolder);
        if (extractedImages.Length == 0)
            throw new InvalidOperationException("No images were extracted to the images folder.");

        // Optional: indicate success (no interactive prompts required).
        Console.WriteLine("Conversion completed successfully.");
        Console.WriteLine($"Markdown file: {Path.GetFullPath(outputMd)}");
        Console.WriteLine($"Extracted images count: {extractedImages.Length}");
    }
}
