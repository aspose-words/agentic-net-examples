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
        string inputDocxPath = "input.docx";
        string outputMarkdownPath = "output.md";
        string imagesFolder = "MarkdownImages";
        string tempImagePath = "sample.png";

        // Ensure the images folder exists.
        Directory.CreateDirectory(imagesFolder);

        // -----------------------------------------------------------------
        // Create a sample DOCX document with some text and an embedded image.
        // -----------------------------------------------------------------
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("This is a sample document containing an image.");

        // Create a simple bitmap using Aspose.Drawing and save it to a temporary file.
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Blue);
            }
            bitmap.Save(tempImagePath, ImageFormat.Png);
        }

        // Insert the temporary image into the document.
        builder.InsertImage(tempImagePath);

        // Save the sample document as DOCX.
        sampleDoc.Save(inputDocxPath, SaveFormat.Docx);

        // Clean up the temporary image file used for insertion.
        if (File.Exists(tempImagePath))
        {
            File.Delete(tempImagePath);
        }

        // -----------------------------------------------------------------
        // Load the DOCX file and convert it to Markdown, extracting images.
        // -----------------------------------------------------------------
        Document doc = new Document(inputDocxPath);
        MarkdownSaveOptions mdOptions = new MarkdownSaveOptions
        {
            ImagesFolder = imagesFolder,
            ImagesFolderAlias = imagesFolder
        };
        doc.Save(outputMarkdownPath, mdOptions);

        // -----------------------------------------------------------------
        // Validation: ensure Markdown file and extracted images were created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputMarkdownPath))
        {
            throw new InvalidOperationException("The Markdown output file was not created.");
        }

        string[] extractedImages = Directory.GetFiles(imagesFolder);
        if (extractedImages.Length == 0)
        {
            throw new InvalidOperationException("No images were extracted to the specified folder.");
        }

        // Optionally, clean up the input DOCX file after conversion.
        // File.Delete(inputDocxPath);
    }
}
