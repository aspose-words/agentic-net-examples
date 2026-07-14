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
        // Define deterministic file names.
        const string gifPath = "sample.gif";
        const string pngPath = "sample.png";
        const string inputDocPath = "input.docx";
        const string outputDocPath = "output.docx";

        // -----------------------------------------------------------------
        // 1. Create a sample image and save it as both GIF and PNG formats.
        // -----------------------------------------------------------------
        const int width = 200;
        const int height = 200;

        // Create a bitmap and draw simple content.
        Bitmap bitmap = new Bitmap(width, height);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.LightBlue);
        // Draw a red ellipse for visual distinction.
        graphics.DrawEllipse(new Pen(Color.Red, 5), 20, 20, width - 40, height - 40);
        // Save as GIF.
        bitmap.Save(gifPath, ImageFormat.Gif);
        // Save as PNG (the replacement image).
        bitmap.Save(pngPath, ImageFormat.Png);
        // Clean up drawing resources.
        graphics.Dispose();
        bitmap.Dispose();

        // --------------------------------------------------------------
        // 2. Create a Word document and insert the GIF image into it.
        // --------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document containing a GIF image:");
        // Insert the GIF image.
        Shape gifShape = builder.InsertImage(gifPath);
        // Ensure the shape indeed holds a GIF.
        if (gifShape.ImageData.ImageType != ImageType.Gif)
            throw new InvalidOperationException("The inserted image is not a GIF as expected.");

        // Save the original document.
        doc.Save(inputDocPath);

        // --------------------------------------------------------------
        // 3. Load the document and replace all GIF images with PNGs.
        // --------------------------------------------------------------
        Document loadedDoc = new Document(inputDocPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int replacedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Gif)
            {
                // Replace the GIF with the corresponding PNG.
                shape.ImageData.SetImage(pngPath);
                replacedCount++;
            }
        }

        // Validate that at least one image was replaced.
        if (replacedCount == 0)
            throw new InvalidOperationException("No GIF images were found to replace.");

        // --------------------------------------------------------------
        // 4. Save the modified document.
        // --------------------------------------------------------------
        loadedDoc.Save(outputDocPath);

        // Simple verification that the output file exists.
        if (!File.Exists(outputDocPath))
            throw new FileNotFoundException("The output document was not created.", outputDocPath);

        // Optional: confirm that all images in the output are PNG.
        Document verifyDoc = new Document(outputDocPath);
        foreach (Shape shape in verifyDoc.GetChildNodes(NodeType.Shape, true).OfType<Shape>())
        {
            if (shape.HasImage && shape.ImageData.ImageType != ImageType.Png)
                throw new InvalidOperationException("An image was not converted to PNG as expected.");
        }

        // The program finishes without requiring user interaction.
    }
}
