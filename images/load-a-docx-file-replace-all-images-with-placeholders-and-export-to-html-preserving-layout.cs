using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Drawing; // Aspose.Drawing provides Bitmap, Graphics, Color

public class Program
{
    public static void Main()
    {
        // Define working directory and file paths.
        string workDir = Directory.GetCurrentDirectory();
        string sampleImagePath = Path.Combine(workDir, "sample.png");
        string placeholderImagePath = Path.Combine(workDir, "placeholder.png");
        string inputDocPath = Path.Combine(workDir, "input.docx");
        string outputHtmlPath = Path.Combine(workDir, "output.html");
        string imagesFolder = Path.Combine(workDir, "html_images");

        // -------------------------------------------------
        // 1. Create a sample image to be inserted into the DOCX.
        // -------------------------------------------------
        using (Bitmap bmp = new Bitmap(200, 150))
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.LightBlue);
            }
            bmp.Save(sampleImagePath);
        }

        // -------------------------------------------------
        // 2. Create a placeholder image that will replace all originals.
        // -------------------------------------------------
        using (Bitmap bmp = new Bitmap(200, 150))
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.LightGray);
            }
            bmp.Save(placeholderImagePath);
        }

        // -------------------------------------------------
        // 3. Build a sample DOCX containing a few images.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Document with images:");
        builder.InsertImage(sampleImagePath);
        builder.Writeln();
        builder.InsertImage(sampleImagePath);
        builder.Writeln();
        builder.InsertImage(sampleImagePath);

        doc.Save(inputDocPath);

        // -------------------------------------------------
        // 4. Load the DOCX, replace each image with the placeholder.
        // -------------------------------------------------
        Document loadedDoc = new Document(inputDocPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Replace the image data with the placeholder image.
                shape.ImageData.SetImage(placeholderImagePath);
            }
        }

        // -------------------------------------------------
        // 5. Save the modified document to HTML, preserving layout.
        // -------------------------------------------------
        if (Directory.Exists(imagesFolder))
            Directory.Delete(imagesFolder, true);
        Directory.CreateDirectory(imagesFolder);

        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            ImagesFolder = imagesFolder,
            ExportImagesAsBase64 = false, // keep images as separate files
            ScaleImageToShapeSize = true   // preserve layout scaling
        };

        loadedDoc.Save(outputHtmlPath, htmlOptions);

        // -------------------------------------------------
        // 6. Simple validation.
        // -------------------------------------------------
        if (!File.Exists(outputHtmlPath))
            throw new InvalidOperationException("HTML output file was not created.");

        if (Directory.GetFiles(imagesFolder).Length == 0)
            throw new InvalidOperationException("No images were saved during HTML export.");

        // Example completed successfully.
    }
}
