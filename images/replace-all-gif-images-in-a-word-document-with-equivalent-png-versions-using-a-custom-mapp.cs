using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // Create a sample GIF image (100x100, light blue background)
        // -----------------------------------------------------------------
        string gifPath = "sample1.gif";
        using (Aspose.Drawing.Bitmap gifBitmap = new Aspose.Drawing.Bitmap(100, 100))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(gifBitmap))
            {
                g.Clear(Aspose.Drawing.Color.LightBlue);
            }
            gifBitmap.Save(gifPath, ImageFormat.Gif);
        }

        // -----------------------------------------------------------------
        // Create the equivalent PNG image (same size, same background)
        // -----------------------------------------------------------------
        string pngPath = "sample1.png";
        using (Aspose.Drawing.Bitmap pngBitmap = new Aspose.Drawing.Bitmap(100, 100))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(pngBitmap))
            {
                g.Clear(Aspose.Drawing.Color.LightBlue);
            }
            pngBitmap.Save(pngPath, ImageFormat.Png);
        }

        // -----------------------------------------------------------------
        // Build a Word document that contains the GIF image
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(gifPath);
        string inputDocPath = "input.docx";
        doc.Save(inputDocPath);

        // -----------------------------------------------------------------
        // Load the document for processing
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(inputDocPath);

        // Mapping from original GIF file name to replacement PNG path.
        // In this simple example we replace any GIF with the prepared PNG.
        var gifToPngMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { Path.GetFileName(gifPath), pngPath }
        };

        int replacedCount = 0;

        // Iterate over all Shape nodes that contain images.
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            // Replace only GIF images.
            if (shape.ImageData.ImageType == ImageType.Gif)
            {
                // Since ImageData does not expose the original file name,
                // we replace any GIF image using the mapping dictionary.
                // If a specific mapping is required, additional logic can be added here.
                string key = Path.GetFileName(gifPath); // known key for this demo
                if (gifToPngMap.TryGetValue(key, out string newPngPath))
                {
                    shape.ImageData.SetImage(newPngPath);
                    replacedCount++;
                }
            }
        }

        // Ensure at least one replacement happened.
        if (replacedCount == 0)
            throw new InvalidOperationException("No GIF images were replaced.");

        // Save the modified document.
        string outputDocPath = "output.docx";
        loadedDoc.Save(outputDocPath);

        // Validate that the output file was created.
        if (!File.Exists(outputDocPath))
            throw new FileNotFoundException("The output document was not created.", outputDocPath);
    }
}
