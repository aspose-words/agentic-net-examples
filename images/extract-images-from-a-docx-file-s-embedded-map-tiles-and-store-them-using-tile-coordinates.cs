using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Loading;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Define folders for artifacts and temporary images.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        string imagesDir = Path.Combine(artifactsDir, "Images");
        Directory.CreateDirectory(artifactsDir);
        Directory.CreateDirectory(imagesDir);

        // -----------------------------------------------------------------
        // 1. Create sample map‑tile images and insert them into a DOCX file.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a 2×2 grid of tiles (you can change the size as needed).
        for (int x = 0; x < 2; x++)
        {
            for (int y = 0; y < 2; y++)
            {
                // Create a deterministic bitmap for the tile.
                using (Bitmap bitmap = new Bitmap(100, 100))
                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    // Fill with a color that depends on the coordinates.
                    int r = (x * 127) % 256;
                    int g = (y * 127) % 256;
                    int b = ((x + y) * 63) % 256;
                    graphics.Clear(Color.FromArgb(r, g, b));

                    // Save the bitmap to a file so it can be inserted.
                    string tileFileName = $"tile_{x}_{y}.png";
                    string tilePath = Path.Combine(imagesDir, tileFileName);
                    bitmap.Save(tilePath);
                }

                // Insert the image into the document.
                string imagePath = Path.Combine(imagesDir, $"tile_{x}_{y}.png");
                Shape shape = builder.InsertImage(imagePath);

                // Store the tile coordinates in the shape's Title property.
                shape.Title = $"tile_{x}_{y}";
            }
        }

        // Save the document containing the map tiles.
        string docPath = Path.Combine(artifactsDir, "MapTiles.docx");
        doc.Save(docPath);

        // ---------------------------------------------------------------
        // 2. Load the document and extract each tile image using its title.
        // ---------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        var shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                                  .OfType<Shape>()
                                  .Where(s => s.HasImage);

        int extractedCount = 0;

        foreach (Shape shape in shapeNodes)
        {
            // The Title holds the original tile coordinates (e.g., "tile_0_1").
            if (string.IsNullOrEmpty(shape.Title))
                continue; // Skip shapes without a title.

            // Determine the file extension based on the image type stored in the shape.
            string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

            // Build the output file name using the coordinates from the title.
            string outputFileName = $"{shape.Title}{extension}";
            string outputPath = Path.Combine(artifactsDir, outputFileName);

            // Save the image data to the file system.
            shape.ImageData.Save(outputPath);
            extractedCount++;
        }

        // Validate that at least one image was extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // Optional: write a short confirmation to the console.
        Console.WriteLine($"Extracted {extractedCount} tile image(s) to \"{artifactsDir}\".");
    }
}
