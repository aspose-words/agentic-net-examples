using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class ExtractMapTileImages
{
    public static void Main()
    {
        // Prepare folders.
        string baseDir = Directory.GetCurrentDirectory();
        string artifactsDir = Path.Combine(baseDir, "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string imagesDir = Path.Combine(artifactsDir, "InputImages");
        Directory.CreateDirectory(imagesDir);
        string outputDir = Path.Combine(artifactsDir, "ExtractedTiles");
        Directory.CreateDirectory(outputDir);

        // Create sample tile images (2x2 grid).
        int tileSize = 100;
        for (int x = 0; x < 2; x++)
        {
            for (int y = 0; y < 2; y++)
            {
                string fileName = Path.Combine(imagesDir, $"tile_{x}_{y}.png");
                using (Bitmap bitmap = new Bitmap(tileSize, tileSize))
                using (Graphics g = Graphics.FromImage(bitmap))
                {
                    // Deterministic color based on coordinates.
                    int r = (x * 127) % 256;
                    int gCol = (y * 127) % 256;
                    int b = ((x + y) * 63) % 256;
                    g.Clear(Color.FromArgb(r, gCol, b));
                    bitmap.Save(fileName, ImageFormat.Png);
                }
            }
        }

        // Create a document and insert the tile images, storing coordinates in the shape title.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int x = 0; x < 2; x++)
        {
            for (int y = 0; y < 2; y++)
            {
                string imagePath = Path.Combine(imagesDir, $"tile_{x}_{y}.png");
                Shape shape = builder.InsertImage(imagePath);
                shape.Title = $"tile_{x}_{y}"; // Store tile coordinates.
                builder.InsertBreak(BreakType.LineBreak);
            }
        }

        string docPath = Path.Combine(artifactsDir, "MapTiles.docx");
        doc.Save(docPath);

        // Load the document and extract images using the stored tile coordinates.
        Document loadedDoc = new Document(docPath);
        var shapes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                              .Cast<Shape>()
                              .Where(s => s.HasImage);

        int extractedCount = 0;
        foreach (var shape in shapes)
        {
            // Retrieve tile coordinates from the shape title.
            string title = shape.Title ?? "tile_unknown";
            string[] parts = title.Split('_');
            if (parts.Length >= 3 &&
                int.TryParse(parts[1], out int tileX) &&
                int.TryParse(parts[2], out int tileY))
            {
                // Determine file extension based on the image type.
                string extension = Aspose.Words.FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string outputFile = Path.Combine(outputDir, $"tile_{tileX}_{tileY}{extension}");
                shape.ImageData.Save(outputFile);
                extractedCount++;
            }
        }

        // Validation: ensure at least one image was extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No tile images were extracted from the document.");

        Console.WriteLine($"Extracted {extractedCount} tile image(s) to: {outputDir}");
    }
}
