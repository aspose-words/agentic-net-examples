using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Define file names and folders.
        const string sampleImagePath = "sample.png";
        const string placeholderImagePath = "placeholder.png";
        const string sourceDocPath = "input.docx";
        const string outputHtmlPath = "output.html";
        const string imagesFolder = "html_images";

        // -------------------------------------------------
        // 1. Create a deterministic sample image (100x100).
        // -------------------------------------------------
        CreateSampleImage(sampleImagePath, 100, 100, Aspose.Drawing.Color.LightBlue);

        // -------------------------------------------------
        // 2. Build a DOCX that contains a few images.
        // -------------------------------------------------
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert three identical images.
        for (int i = 0; i < 3; i++)
        {
            builder.InsertImage(sampleImagePath);
            builder.Writeln(); // Add a line break between images.
        }

        // Save the source document.
        doc.Save(sourceDocPath);

        // -------------------------------------------------
        // 3. Create a placeholder image that will replace all originals.
        // -------------------------------------------------
        CreateSampleImage(placeholderImagePath, 100, 100, Aspose.Drawing.Color.LightGray);

        // -------------------------------------------------
        // 4. Load the document, replace each image with the placeholder.
        // -------------------------------------------------
        var loadedDoc = new Document(sourceDocPath);
        var shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        foreach (Shape shape in shapes)
        {
            if (shape.HasImage)
            {
                // Replace the image data with the placeholder image.
                shape.ImageData.SetImage(placeholderImagePath);
            }
        }

        // -------------------------------------------------
        // 5. Save the modified document as HTML, preserving layout.
        // -------------------------------------------------
        // Ensure the images folder exists; Aspose.Words will create it if missing,
        // but we create it explicitly for validation purposes.
        if (Directory.Exists(imagesFolder))
            Directory.Delete(imagesFolder, true);
        Directory.CreateDirectory(imagesFolder);

        var htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            ImagesFolder = imagesFolder,
            ExportImagesAsBase64 = false, // Save images as separate files.
            ScaleImageToShapeSize = true   // Preserve layout scaling.
        };

        loadedDoc.Save(outputHtmlPath, htmlOptions);

        // -------------------------------------------------
        // 6. Validation – ensure output files were created.
        // -------------------------------------------------
        if (!File.Exists(outputHtmlPath))
            throw new InvalidOperationException("HTML output file was not created.");

        if (!Directory.Exists(imagesFolder))
            throw new InvalidOperationException("Images folder was not created.");

        var imageFiles = Directory.GetFiles(imagesFolder);
        if (imageFiles.Length == 0)
            throw new InvalidOperationException("No images were saved during HTML export.");

        // Optional: clean up temporary files (comment out if inspection is needed).
        //File.Delete(sampleImagePath);
        //File.Delete(placeholderImagePath);
        //File.Delete(sourceDocPath);
        //File.Delete(outputHtmlPath);
        //Directory.Delete(imagesFolder, true);
    }

    // Helper method to create a deterministic bitmap and save it to a file.
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color backgroundColor)
    {
        var bitmap = new Bitmap(width, height);
        var graphics = Graphics.FromImage(bitmap);
        graphics.Clear(backgroundColor);
        // Deterministic drawing can be added here if needed.
        graphics.Dispose();
        bitmap.Save(filePath);
        bitmap.Dispose();
    }
}
