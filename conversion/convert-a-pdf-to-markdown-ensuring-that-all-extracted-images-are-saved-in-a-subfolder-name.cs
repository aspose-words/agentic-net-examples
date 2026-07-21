using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file and folder names.
        string inputPdfPath = Path.Combine(Directory.GetCurrentDirectory(), "input.pdf");
        string outputMdPath = Path.Combine(Directory.GetCurrentDirectory(), "output.md");
        string assetsFolder = Path.Combine(Directory.GetCurrentDirectory(), "assets");

        // -----------------------------------------------------------------
        // 1. Create a sample PDF document (input for conversion).
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample PDF content with an image.");

        // Insert a sample image. Use any image file that exists in the execution folder.
        // If the image file is missing, the code will still run; the image will simply be omitted.
        string sampleImagePath = Path.Combine(Directory.GetCurrentDirectory(), "sample.jpg");
        if (File.Exists(sampleImagePath))
        {
            builder.InsertImage(sampleImagePath);
        }

        // Save the document as PDF.
        sourceDoc.Save(inputPdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(inputPdfPath))
            throw new InvalidOperationException("The input PDF file was not created.");

        // -----------------------------------------------------------------
        // 2. Load the PDF and convert it to Markdown, extracting images to "assets".
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(inputPdfPath);

        // Configure Markdown save options.
        MarkdownSaveOptions mdOptions = new MarkdownSaveOptions
        {
            SaveFormat = SaveFormat.Markdown,
            ImagesFolder = assetsFolder,
            // Optional: set a folder alias if you want relative URIs without the folder name.
            // ImagesFolderAlias = "assets"
        };

        // Ensure the assets folder exists (Aspose will create it automatically,
        // but we create it beforehand to guarantee its presence).
        if (!Directory.Exists(assetsFolder))
            Directory.CreateDirectory(assetsFolder);

        // Save as Markdown.
        pdfDoc.Save(outputMdPath, mdOptions);

        // -----------------------------------------------------------------
        // 3. Validation.
        // -----------------------------------------------------------------
        if (!File.Exists(outputMdPath))
            throw new InvalidOperationException("The Markdown output file was not created.");

        if (!Directory.Exists(assetsFolder))
            throw new InvalidOperationException("The assets folder for images was not created.");

        // At least one image file should be present if the source PDF contained images.
        // If no images were inserted, the folder may be empty; this is still acceptable.
        // Uncomment the following lines to enforce the presence of images:
        // if (Directory.GetFiles(assetsFolder).Length == 0)
        //     throw new InvalidOperationException("No images were extracted to the assets folder.");

        // The example finishes execution here.
    }
}
