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
        // Create a temporary working directory.
        string workDir = Path.Combine(Path.GetTempPath(), "AsposeExample_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(workDir);

        // Create a sample image using Aspose.Drawing.
        string imagePath = Path.Combine(workDir, "sample.png");
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Blue);
            }
            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // Build a simple document that includes the image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document with an image:");
        builder.InsertImage(imagePath);

        // Save the document as PDF (the source for conversion).
        string pdfPath = Path.Combine(workDir, "input.pdf");
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF back into Aspose.Words.
        Document pdfDoc = new Document(pdfPath);

        // Prepare the folder where extracted images will be saved.
        string imagesFolder = Path.Combine(workDir, "images");
        Directory.CreateDirectory(imagesFolder);
        string imagesFolderAlias = "./images";

        // Configure Markdown save options to extract images.
        MarkdownSaveOptions mdOptions = new MarkdownSaveOptions
        {
            ImagesFolder = imagesFolder,
            ImagesFolderAlias = imagesFolderAlias
        };

        // Save the PDF as Markdown.
        string mdPath = Path.Combine(workDir, "output.md");
        pdfDoc.Save(mdPath, mdOptions);

        // Validation.
        if (!File.Exists(mdPath))
            throw new InvalidOperationException("Markdown file was not created.");

        if (!Directory.Exists(imagesFolder) || Directory.GetFiles(imagesFolder).Length == 0)
            throw new InvalidOperationException("No images were extracted to the images folder.");

        // Output the locations for verification (optional).
        Console.WriteLine("Conversion completed successfully.");
        Console.WriteLine($"Working directory: {workDir}");
        Console.WriteLine($"Markdown file: {mdPath}");
        Console.WriteLine($"Images folder: {imagesFolder}");
    }
}
