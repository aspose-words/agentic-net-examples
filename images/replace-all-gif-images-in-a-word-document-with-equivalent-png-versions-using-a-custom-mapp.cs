using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;          // Aspose.Drawing for Bitmap, Graphics, Color, etc.

public class Program
{
    public static void Main()
    {
        // Define deterministic file names (non‑constant because they are built at runtime).
        string artifactsDir = "Artifacts";
        string gifPath = Path.Combine(artifactsDir, "sample.gif");
        string pngPath = Path.Combine(artifactsDir, "sample.png");
        string inputDocPath = Path.Combine(artifactsDir, "input.docx");
        string outputDocPath = Path.Combine(artifactsDir, "output.docx");

        // Ensure the output folder exists.
        Directory.CreateDirectory(artifactsDir);

        // -------------------------------------------------
        // 1. Create a sample GIF image.
        // -------------------------------------------------
        using (Bitmap gifBitmap = new Bitmap(100, 100))
        {
            using (Graphics g = Graphics.FromImage(gifBitmap))
            {
                g.Clear(Color.White);
                g.FillRectangle(new SolidBrush(Color.Red), 10, 10, 80, 80);
            }
            // Save as GIF.
            gifBitmap.Save(gifPath, Aspose.Drawing.Imaging.ImageFormat.Gif);
        }

        // -------------------------------------------------
        // 2. Create a corresponding PNG image (same visual content).
        // -------------------------------------------------
        using (Bitmap pngBitmap = new Bitmap(100, 100))
        {
            using (Graphics g = Graphics.FromImage(pngBitmap))
            {
                g.Clear(Color.White);
                g.FillRectangle(new SolidBrush(Color.Red), 10, 10, 80, 80);
            }
            // Save as PNG.
            pngBitmap.Save(pngPath, Aspose.Drawing.Imaging.ImageFormat.Png);
        }

        // -------------------------------------------------
        // 3. Create a Word document that contains the GIF image.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document with a GIF image:");
        // Insert the GIF image.
        Shape gifShape = builder.InsertImage(gifPath);
        // Verify that the inserted image is indeed a GIF.
        if (!gifShape.HasImage || gifShape.ImageData.ImageType != ImageType.Gif)
            throw new InvalidOperationException("Inserted image is not a GIF as expected.");

        // Save the original document.
        doc.Save(inputDocPath);

        // -------------------------------------------------
        // 4. Load the document and replace GIF images with PNG equivalents.
        // -------------------------------------------------
        Document loadedDoc = new Document(inputDocPath);

        // Custom mapping: GIF file name -> PNG file name.
        // In this simple example we map any GIF to the prepared PNG.
        string mappingGifToPng = pngPath;

        // Iterate over all Shape nodes.
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int replacedCount = 0;
        foreach (Shape shape in shapes)
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Gif)
            {
                // Replace the image data with the PNG version.
                shape.ImageData.SetImage(mappingGifToPng);
                replacedCount++;
            }
        }

        // Validation: ensure at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No GIF images were found to replace.");

        // Ensure no GIF images remain.
        foreach (Shape shape in shapes)
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Gif)
                throw new InvalidOperationException("A GIF image still remains after replacement.");
        }

        // Save the modified document.
        loadedDoc.Save(outputDocPath);

        // Final validation: output file must exist.
        if (!File.Exists(outputDocPath))
            throw new FileNotFoundException("The output document was not created.", outputDocPath);

        // Indicate success (no interactive prompts required).
        Console.WriteLine("GIF images successfully replaced with PNG equivalents.");
    }
}
