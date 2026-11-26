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
        // Define file and folder paths
        string inputDocxPath = "input.docx";
        string outputMarkdownPath = "output.md";
        string imagesFolder = "output_images";
        string tempImagePath = "temp_image.png";

        // Ensure the images folder exists
        Directory.CreateDirectory(imagesFolder);

        // Create a sample DOCX document with text and an image
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample document with an image.");

        // Create a simple bitmap using Aspose.Drawing
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Blue);
            }

            // Save the bitmap to a temporary PNG file
            bitmap.Save(tempImagePath, ImageFormat.Png);
        }

        // Insert the temporary image into the document
        builder.InsertImage(tempImagePath);

        // Save the source document as DOCX
        sourceDoc.Save(inputDocxPath, SaveFormat.Docx);

        // Load the DOCX document
        Document doc = new Document(inputDocxPath);

        // Configure Markdown save options with a custom images folder
        MarkdownSaveOptions mdOptions = new MarkdownSaveOptions
        {
            ImagesFolder = imagesFolder,
            ImagesFolderAlias = imagesFolder
        };

        // Save the document as Markdown
        doc.Save(outputMarkdownPath, mdOptions);

        // Validate that the Markdown file was created
        if (!File.Exists(outputMarkdownPath))
            throw new InvalidOperationException("The Markdown file was not created.");

        // Validate that at least one image was extracted to the images folder
        string[] extractedImages = Directory.GetFiles(imagesFolder);
        if (extractedImages.Length == 0)
            throw new InvalidOperationException("No images were extracted to the images folder.");

        // Clean up the temporary image file used for insertion
        if (File.Exists(tempImagePath))
            File.Delete(tempImagePath);
    }
}
