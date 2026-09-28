using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class ReplaceGifWithPng
{
    public static void Main()
    {
        // Create sample GIF image.
        const string gifPath = "sample.gif";
        const string pngPath = "sample.png";

        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.LightBlue);
                // Draw a simple rectangle.
                g.FillRectangle(new SolidBrush(Aspose.Drawing.Color.Red), 10, 10, 80, 80);
            }

            // Save as GIF.
            bitmap.Save(gifPath, ImageFormat.Gif);
            // Save as PNG (the replacement image).
            bitmap.Save(pngPath, ImageFormat.Png);
        }

        // Create a Word document containing the GIF image.
        const string inputDoc = "input.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(gifPath);
        doc.Save(inputDoc);

        // Load the document for processing.
        Document loadedDoc = new Document(inputDoc);

        // Mapping from original image type to replacement image file.
        var replacementMap = new Dictionary<ImageType, string>
        {
            { ImageType.Gif, pngPath }
        };

        int replacedCount = 0;

        // Iterate over all Shape nodes and replace GIF images.
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes)
        {
            if (shape.HasImage && replacementMap.TryGetValue(shape.ImageData.ImageType, out string newImagePath))
            {
                shape.ImageData.SetImage(newImagePath);
                replacedCount++;
            }
        }

        // Validate that at least one image was replaced.
        if (replacedCount == 0)
            throw new InvalidOperationException("No GIF images were found to replace.");

        // Save the updated document.
        const string outputDoc = "output.docx";
        loadedDoc.Save(outputDoc);

        // Validate output file existence.
        if (!File.Exists(outputDoc))
            throw new FileNotFoundException("The output document was not created.", outputDoc);
    }
}
