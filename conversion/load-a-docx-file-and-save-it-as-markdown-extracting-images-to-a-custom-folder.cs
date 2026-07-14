using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;   // Needed for ImageFormat

public class Program
{
    public static void Main()
    {
        // Define file and folder names.
        const string inputDocx = "input.docx";
        const string outputMd = "output.md";
        const string imagesFolder = "Images";
        const string sampleImagePath = "sample.png";

        // Ensure a clean environment.
        if (File.Exists(inputDocx)) File.Delete(inputDocx);
        if (File.Exists(outputMd)) File.Delete(outputMd);
        if (Directory.Exists(imagesFolder)) Directory.Delete(imagesFolder, true);
        if (File.Exists(sampleImagePath)) File.Delete(sampleImagePath);

        // Create a simple image using Aspose.Drawing.
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            // Draw a solid blue rectangle.
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Blue);
            }

            // Save the bitmap as a PNG file.
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // Create a sample DOCX document that contains text and the image.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Hello World!");
        builder.InsertImage(sampleImagePath);
        sampleDoc.Save(inputDocx, SaveFormat.Docx);

        // Load the DOCX document.
        Document doc = new Document(inputDocx);

        // Prepare Markdown save options with a custom images folder.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            ImagesFolder = imagesFolder,
            SaveFormat = SaveFormat.Markdown   // Explicitly set the format.
        };

        // Ensure the images folder exists.
        Directory.CreateDirectory(imagesFolder);

        // Save the document as Markdown.
        doc.Save(outputMd, saveOptions);

        // Validate that the Markdown file was created.
        if (!File.Exists(outputMd))
            throw new InvalidOperationException("The Markdown output file was not created.");

        // Validate that at least one image was extracted.
        string[] extractedImages = Directory.GetFiles(imagesFolder);
        if (extractedImages.Length == 0)
            throw new InvalidOperationException("No images were extracted to the specified folder.");

        // Indicate success.
        Console.WriteLine("Conversion completed successfully.");
        Console.WriteLine($"Markdown file: {Path.GetFullPath(outputMd)}");
        Console.WriteLine($"Extracted images folder: {Path.GetFullPath(imagesFolder)}");
    }
}
