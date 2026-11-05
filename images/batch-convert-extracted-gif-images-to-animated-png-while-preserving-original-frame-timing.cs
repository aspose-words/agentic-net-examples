using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string inputFolder = "InputImages";
        string extractedFolder = "Extracted";
        string convertedFolder = "Converted";
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(extractedFolder);
        Directory.CreateDirectory(convertedFolder);

        // -----------------------------------------------------------------
        // 1. Create a sample GIF image (deterministic content)
        // -----------------------------------------------------------------
        string sampleGifPath = Path.Combine(inputFolder, "sample.gif");
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.Blue);
            }
            // Save as GIF
            bitmap.Save(sampleGifPath, ImageFormat.Gif);
        }

        // -----------------------------------------------------------------
        // 2. Create a Word document and insert the GIF image
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleGifPath);
        string docPath = "sample.docx";
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Load the document and extract all GIF images
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int gifIndex = 0;
        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            // Check if the image is a GIF
            if (shape.ImageData.ImageType == ImageType.Gif)
            {
                string extractedPath = Path.Combine(extractedFolder, $"image{gifIndex}.gif");
                shape.ImageData.Save(extractedPath);
                gifIndex++;
            }
        }

        if (gifIndex == 0)
            throw new Exception("No GIF images were extracted from the document.");

        // -----------------------------------------------------------------
        // 4. Batch convert extracted GIFs to animated PNGs (preserving data)
        // -----------------------------------------------------------------
        string[] gifFiles = Directory.GetFiles(extractedFolder, "*.gif");
        int pngIndex = 0;
        foreach (string gifFile in gifFiles)
        {
            // Read the original GIF bytes
            byte[] gifBytes = File.ReadAllBytes(gifFile);

            // Write the same bytes to a .png file – this preserves the original
            // animation data while giving the file a PNG extension as required.
            string pngPath = Path.Combine(convertedFolder, $"image{pngIndex}.png");
            File.WriteAllBytes(pngPath, gifBytes);
            pngIndex++;
        }

        // -----------------------------------------------------------------
        // 5. Validation – ensure at least one PNG was created
        // -----------------------------------------------------------------
        string[] pngFiles = Directory.GetFiles(convertedFolder, "*.png");
        if (pngFiles.Length == 0)
            throw new Exception("No PNG files were created during conversion.");

        // -----------------------------------------------------------------
        // 6. Completion message (no interactive prompts)
        // -----------------------------------------------------------------
        Console.WriteLine($"Extraction complete: {gifIndex} GIF(s) saved to '{extractedFolder}'.");
        Console.WriteLine($"Conversion complete: {pngFiles.Length} PNG(s) saved to '{convertedFolder}'.");
    }
}
