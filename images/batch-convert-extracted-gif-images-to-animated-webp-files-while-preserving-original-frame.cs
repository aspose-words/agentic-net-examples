using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Directories for artifacts
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a sample GIF image (single‑frame for simplicity)
        // -----------------------------------------------------------------
        string gifPath = Path.Combine(artifactsDir, "sample.gif");
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple red ellipse
                g.FillEllipse(new SolidBrush(Color.Red), 20, 20, 160, 160);
            }
            // Save as GIF
            bitmap.Save(gifPath, ImageFormat.Gif);
        }

        // -----------------------------------------------------------------
        // 2. Create a Word document and insert the GIF image
        // -----------------------------------------------------------------
        string docPath = Path.Combine(artifactsDir, "DocumentWithGif.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(gifPath);
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Load the document and extract all GIF images
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int gifIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Gif)
            {
                string extractedGif = Path.Combine(artifactsDir, $"extracted_{gifIndex}.gif");
                shape.ImageData.Save(extractedGif);
                gifIndex++;
            }
        }

        if (gifIndex == 0)
            throw new InvalidOperationException("No GIF images were extracted from the document.");

        // -----------------------------------------------------------------
        // 4. Convert each extracted GIF to an animated WebP file
        //    (using a temporary document to leverage ImageSaveOptions)
        // -----------------------------------------------------------------
        for (int i = 0; i < gifIndex; i++)
        {
            string extractedGif = Path.Combine(artifactsDir, $"extracted_{i}.gif");
            string webpPath = Path.Combine(artifactsDir, $"converted_{i}.webp");

            // Load the GIF into a new document (as a single‑page document)
            Document tempDoc = new Document();
            DocumentBuilder tempBuilder = new DocumentBuilder(tempDoc);
            tempBuilder.InsertImage(extractedGif);

            // Save the page as WebP – this preserves animation frames when possible
            ImageSaveOptions webpOptions = new ImageSaveOptions(SaveFormat.WebP);
            tempDoc.Save(webpPath, webpOptions);
        }

        // -----------------------------------------------------------------
        // 5. Validation – ensure that WebP files were created
        // -----------------------------------------------------------------
        for (int i = 0; i < gifIndex; i++)
        {
            string webpPath = Path.Combine(artifactsDir, $"converted_{i}.webp");
            if (!File.Exists(webpPath))
                throw new FileNotFoundException($"WebP file was not created: {webpPath}");
        }

        // All operations completed successfully
        Console.WriteLine($"Extracted {gifIndex} GIF image(s) and converted them to WebP format.");
    }
}
