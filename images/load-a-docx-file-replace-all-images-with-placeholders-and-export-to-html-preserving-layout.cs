using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Define file and folder names
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        string imagesDir = Path.Combine(workDir, "Images");
        string htmlImagesDir = Path.Combine(workDir, "HtmlImages");
        string docPath = Path.Combine(workDir, "input.docx");
        string htmlPath = Path.Combine(workDir, "output.html");
        string sampleImagePath = Path.Combine(imagesDir, "sample.png");
        string placeholderImagePath = Path.Combine(imagesDir, "placeholder.png");

        // Ensure clean workspace
        if (Directory.Exists(workDir))
            Directory.Delete(workDir, true);
        Directory.CreateDirectory(workDir);
        Directory.CreateDirectory(imagesDir);
        Directory.CreateDirectory(htmlImagesDir);

        // -------------------------------------------------
        // 1. Create a sample image (sample.png)
        // -------------------------------------------------
        const int sampleWidth = 200;
        const int sampleHeight = 150;
        using (Bitmap bmp = new Bitmap(sampleWidth, sampleHeight))
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.LightBlue);
                // Simple visual content – a filled ellipse
                g.FillEllipse(Brushes.DarkBlue, 20, 20, sampleWidth - 40, sampleHeight - 40);
            }
            bmp.Save(sampleImagePath);
        }

        // -------------------------------------------------
        // 2. Create a placeholder image (placeholder.png)
        // -------------------------------------------------
        const int placeholderSize = 100;
        using (Bitmap bmp = new Bitmap(placeholderSize, placeholderSize))
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.LightGray);
                // Simple visual content – a red cross
                g.DrawLine(Pens.Red, 0, 0, placeholderSize, placeholderSize);
                g.DrawLine(Pens.Red, placeholderSize, 0, 0, placeholderSize);
            }
            bmp.Save(placeholderImagePath);
        }

        // -------------------------------------------------
        // 3. Create a DOCX document and insert the sample image several times
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document with images:");
        for (int i = 0; i < 3; i++)
        {
            builder.InsertImage(sampleImagePath);
            builder.Writeln(); // add a line break after each image
        }
        doc.Save(docPath);

        // -------------------------------------------------
        // 4. Load the document, replace each image with the placeholder image
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Replace the image data with the placeholder image file
                shape.ImageData.SetImage(placeholderImagePath);
            }
        }

        // -------------------------------------------------
        // 5. Save the modified document to HTML, preserving layout
        // -------------------------------------------------
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            ImagesFolder = htmlImagesDir,
            ExportImagesAsBase64 = false, // keep images as separate files
            ScaleImageToShapeSize = true   // ensure layout matches original shape sizes
        };
        loadedDoc.Save(htmlPath, htmlOptions);

        // -------------------------------------------------
        // 6. Validate that output files were created
        // -------------------------------------------------
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("HTML file was not created.");

        if (!Directory.Exists(htmlImagesDir) || !Directory.GetFiles(htmlImagesDir).Any())
            throw new InvalidOperationException("No images were saved during HTML export.");

        // The example finishes without requiring user interaction.
    }
}
