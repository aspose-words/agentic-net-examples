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
        // Base output directory.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(baseDir);

        // Folder where extracted images will be saved.
        string imagesFolder = Path.Combine(baseDir, "Images");
        Directory.CreateDirectory(imagesFolder);

        // Paths for the temporary files.
        string docxPath = Path.Combine(baseDir, "Sample.docx");
        string markdownPath = Path.Combine(baseDir, "Sample.md");
        string imagePath = Path.Combine(baseDir, "sample.png");

        // Create a simple PNG image using Aspose.Drawing.
        CreateSampleImage(imagePath);

        // Create a DOCX document that contains some text and the image.
        CreateSampleDocx(docxPath, imagePath);

        // Load the DOCX document.
        Document doc = new Document(docxPath);

        // Configure Markdown save options to extract images to the custom folder.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            ImagesFolder = imagesFolder,          // Physical folder for image files.
            ImagesFolderAlias = "Images"          // URI used inside the Markdown file.
        };

        // Save the document as Markdown.
        doc.Save(markdownPath, saveOptions);

        // Validation: ensure the Markdown file exists.
        if (!File.Exists(markdownPath))
            throw new InvalidOperationException("Markdown file was not created.");

        // Validation: ensure at least one image was extracted.
        string[] extractedImages = Directory.GetFiles(imagesFolder);
        if (extractedImages.Length == 0)
            throw new InvalidOperationException("No images were extracted to the Images folder.");
    }

    // Generates a 100x100 red square PNG image using Aspose.Drawing.
    private static void CreateSampleImage(string path)
    {
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            // Obtain a Graphics object that can draw onto the bitmap.
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Red);
            }

            // Save the bitmap as a PNG file.
            bitmap.Save(path, ImageFormat.Png);
        }
    }

    // Creates a DOCX file with a paragraph of text and inserts the provided image.
    private static void CreateSampleDocx(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document that contains an image:");
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }
}
