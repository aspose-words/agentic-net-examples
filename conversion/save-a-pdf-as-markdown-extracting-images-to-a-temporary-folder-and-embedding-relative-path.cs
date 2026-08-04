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
        // Define file and folder paths.
        string inputPdfPath = "input.pdf";
        string markdownPath = "output.md";
        string imagesFolder = "Images";
        string imageFileName = "sample.png";

        // Ensure the images folder exists.
        Directory.CreateDirectory(imagesFolder);

        // Step 1: Create a sample image using Aspose.Drawing.
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            // Obtain a Graphics object for the bitmap.
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill the bitmap with a solid red color.
                graphics.Clear(Color.Red);
            }

            // Save the bitmap as a PNG file.
            bitmap.Save(imageFileName, ImageFormat.Png);
        }

        // Step 2: Create a Word document, insert text and the sample image,
        // then save it as a PDF file (input.pdf).
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document containing an image.");
        builder.InsertImage(imageFileName);
        doc.Save(inputPdfPath, SaveFormat.Pdf);

        // Step 3: Load the PDF document.
        Document pdfDoc = new Document(inputPdfPath);

        // Step 4: Configure MarkdownSaveOptions to extract images to a folder
        // and embed relative paths in the Markdown output.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            ImagesFolder = imagesFolder,
            ImagesFolderAlias = "./Images"
        };

        // Step 5: Save the PDF as a Markdown file using the configured options.
        pdfDoc.Save(markdownPath, saveOptions);

        // Step 6: Validation - ensure the Markdown file and extracted images exist.
        if (!File.Exists(markdownPath))
            throw new InvalidOperationException("Markdown output file was not created.");

        if (!Directory.Exists(imagesFolder))
            throw new InvalidOperationException("Images folder was not created.");

        string[] extractedImages = Directory.GetFiles(imagesFolder);
        if (extractedImages.Length == 0)
            throw new InvalidOperationException("No images were extracted to the images folder.");

        // Output paths for verification.
        Console.WriteLine($"Markdown file created at: {Path.GetFullPath(markdownPath)}");
        Console.WriteLine($"Extracted images are located in: {Path.GetFullPath(imagesFolder)}");
    }
}
