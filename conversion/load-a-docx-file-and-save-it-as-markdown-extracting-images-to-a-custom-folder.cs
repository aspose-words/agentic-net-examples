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
        const string sampleImage = "sample.png";

        // Ensure a clean environment.
        if (File.Exists(inputDocx)) File.Delete(inputDocx);
        if (File.Exists(outputMd)) File.Delete(outputMd);
        if (Directory.Exists(imagesFolder)) Directory.Delete(imagesFolder, true);
        if (File.Exists(sampleImage)) File.Delete(sampleImage);

        // Create a simple PNG image using Aspose.Drawing.
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            // Obtain a Graphics object from the bitmap.
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Blue);
            }
            // Save the bitmap as a PNG file.
            bitmap.Save(sampleImage, ImageFormat.Png);
        }

        // Create a sample DOCX document that contains the image.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample document with an image.");
        builder.InsertImage(sampleImage);
        sourceDoc.Save(inputDocx, SaveFormat.Docx);

        // Load the DOCX document.
        Document doc = new Document(inputDocx);

        // Prepare the images folder.
        Directory.CreateDirectory(imagesFolder);

        // Configure Markdown save options to extract images to the custom folder.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            ImagesFolder = imagesFolder,
            SaveFormat = SaveFormat.Markdown
        };

        // Save the document as Markdown.
        doc.Save(outputMd, saveOptions);

        // Validation: ensure the Markdown file was created.
        if (!File.Exists(outputMd))
            throw new InvalidOperationException("The Markdown output file was not created.");

        // Validation: ensure at least one image was extracted.
        string[] extractedImages = Directory.GetFiles(imagesFolder);
        if (extractedImages.Length == 0)
            throw new InvalidOperationException("No images were extracted to the specified folder.");

        // Clean up the temporary sample image.
        if (File.Exists(sampleImage))
            File.Delete(sampleImage);
    }
}
