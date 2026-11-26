using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeDrawing = Aspose.Drawing;
using AsposeDrawingImaging = Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare deterministic file names.
        const string gifPath = "sample.gif";
        const string pngPath = "sample.png";
        const string inputDocPath = "input.docx";
        const string outputDocPath = "output.docx";

        // Create a sample GIF image.
        using (AsposeDrawing.Bitmap bitmap = new AsposeDrawing.Bitmap(200, 100))
        {
            using (AsposeDrawing.Graphics g = AsposeDrawing.Graphics.FromImage(bitmap))
            {
                g.Clear(AsposeDrawing.Color.White);
                g.DrawRectangle(AsposeDrawing.Pens.Black, 10, 10, 180, 80);
                g.DrawString("GIF Image",
                    new AsposeDrawing.Font("Arial", 12),
                    AsposeDrawing.Brushes.Black,
                    new AsposeDrawing.PointF(30, 40));
            }
            bitmap.Save(gifPath, AsposeDrawingImaging.ImageFormat.Gif);
        }

        // Create the PNG counterpart (same visual content).
        using (AsposeDrawing.Bitmap bitmap = new AsposeDrawing.Bitmap(200, 100))
        {
            using (AsposeDrawing.Graphics g = AsposeDrawing.Graphics.FromImage(bitmap))
            {
                g.Clear(AsposeDrawing.Color.White);
                g.DrawRectangle(AsposeDrawing.Pens.Black, 10, 10, 180, 80);
                g.DrawString("PNG Image",
                    new AsposeDrawing.Font("Arial", 12),
                    AsposeDrawing.Brushes.Black,
                    new AsposeDrawing.PointF(30, 40));
            }
            bitmap.Save(pngPath, AsposeDrawingImaging.ImageFormat.Png);
        }

        // Verify that the images were created.
        if (!File.Exists(gifPath) || !File.Exists(pngPath))
            throw new Exception("Failed to create sample images.");

        // Create a new document and insert the GIF image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(gifPath);
        doc.Save(inputDocPath);

        // Verify input document creation.
        if (!File.Exists(inputDocPath))
            throw new Exception("Failed to create input document.");

        // Load the document for processing.
        Document loadedDoc = new Document(inputDocPath);

        // Custom mapping from GIF file name to PNG file name.
        Dictionary<string, string> gifToPngMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { gifPath, pngPath }
        };

        int replacedCount = 0;

        // Iterate over all Shape nodes that may contain images.
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            // Check if the image is a GIF.
            if (shape.ImageData.ImageType == ImageType.Gif)
            {
                // In this example we know the source key.
                string sourceKey = gifPath;

                if (gifToPngMap.TryGetValue(sourceKey, out string newImagePath) && File.Exists(newImagePath))
                {
                    shape.ImageData.SetImage(newImagePath);
                    replacedCount++;
                }
            }
        }

        // Ensure at least one replacement occurred.
        if (replacedCount == 0)
            throw new Exception("No GIF images were found and replaced.");

        // Save the modified document.
        loadedDoc.Save(outputDocPath);

        // Validate output.
        if (!File.Exists(outputDocPath))
            throw new Exception("Failed to save the output document.");

        // Optional cleanup (commented out).
        // File.Delete(gifPath);
        // File.Delete(pngPath);
        // File.Delete(inputDocPath);
    }
}
