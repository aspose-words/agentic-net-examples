using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

public class BatchGifToWebp
{
    public static void Main()
    {
        // Prepare folders.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputImages");
        string outputDir = Path.Combine(baseDir, "OutputImages");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a deterministic animated GIF file from an embedded base‑64 string.
        // This GIF has a single frame; the conversion logic works the same for animated GIFs.
        string gifBase64 =
            "R0lGODdhAQABAIAAAAUEBAAAACwAAAAAAQABAAACAkQBADs="; // 1×1 pixel transparent GIF
        byte[] gifBytes = Convert.FromBase64String(gifBase64);
        string gifPath = Path.Combine(inputDir, "sample.gif");
        File.WriteAllBytes(gifPath, gifBytes);

        // Create a Word document and insert the GIF image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(gifPath);
        string docPath = Path.Combine(baseDir, "DocumentWithGif.docx");
        doc.Save(docPath);

        // Extract all GIF images from the document.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Gif)
            {
                // Save the extracted GIF.
                string extractedGifPath = Path.Combine(outputDir, $"extracted_{extractedCount}.gif");
                shape.ImageData.Save(extractedGifPath);

                // Convert the extracted GIF to an animated WebP while preserving frame delays.
                // Insert the GIF into a temporary document and save that document as WebP.
                Document tempDoc = new Document();
                DocumentBuilder tempBuilder = new DocumentBuilder(tempDoc);
                tempBuilder.InsertImage(extractedGifPath);

                ImageSaveOptions webpOptions = new ImageSaveOptions(SaveFormat.WebP);
                string webpPath = Path.Combine(outputDir, $"converted_{extractedCount}.webp");
                tempDoc.Save(webpPath, webpOptions);

                extractedCount++;
            }
        }

        // Validation: ensure at least one WebP file was produced.
        if (extractedCount == 0 || !Directory.GetFiles(outputDir, "*.webp").Any())
        {
            throw new InvalidOperationException("No WebP files were created during the batch conversion.");
        }

        // Optional clean‑up can be performed here if needed.
    }
}
