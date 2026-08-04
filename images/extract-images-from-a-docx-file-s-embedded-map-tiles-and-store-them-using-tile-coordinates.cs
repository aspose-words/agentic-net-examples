using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class ExtractMapTileImages
{
    public static void Main()
    {
        // Prepare folders.
        string baseDir = Directory.GetCurrentDirectory();
        string imagesInputDir = Path.Combine(baseDir, "InputTiles");
        string imagesOutputDir = Path.Combine(baseDir, "ExtractedTiles");
        Directory.CreateDirectory(imagesInputDir);
        Directory.CreateDirectory(imagesOutputDir);

        // Create sample tile images (3x2 grid) using Aspose.Drawing.
        int tileWidth = 100;
        int tileHeight = 100;
        for (int x = 0; x < 3; x++)
        {
            for (int y = 0; y < 2; y++)
            {
                string fileName = $"tile_{x}_{y}.png";
                string filePath = Path.Combine(imagesInputDir, fileName);

                using (Bitmap bitmap = new Bitmap(tileWidth, tileHeight))
                using (Graphics g = Graphics.FromImage(bitmap))
                {
                    // Fill background with a color based on coordinates.
                    int r = (x * 80) % 256;
                    int gCol = (y * 120) % 256;
                    int b = ((x + y) * 60) % 256;
                    g.Clear(Color.FromArgb(r, gCol, b));

                    // Optionally draw the coordinates (not required for extraction).
                    // Save the bitmap.
                    bitmap.Save(filePath, ImageFormat.Png);
                }
            }
        }

        // Create a DOCX and insert the tile images, storing coordinates in AlternativeText.
        string docPath = Path.Combine(baseDir, "MapTiles.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        foreach (string tileFile in Directory.GetFiles(imagesInputDir, "*.png"))
        {
            // Insert image.
            Shape shape = builder.InsertImage(tileFile);
            // Store tile coordinates (extracted from file name) in AlternativeText.
            string fileName = Path.GetFileNameWithoutExtension(tileFile); // e.g., tile_0_1
            shape.AlternativeText = fileName.Replace("tile_", ""); // "0_1"
        }

        // Save the document.
        doc.Save(docPath);

        // Load the document and extract images using tile coordinates.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Determine tile coordinates from AlternativeText; fallback to index if missing.
            string coordPart = shape.AlternativeText;
            if (string.IsNullOrWhiteSpace(coordPart))
                coordPart = $"idx_{extractedCount}";

            // Build output file name with proper extension.
            string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
            string outFileName = $"tile_{coordPart}{extension}";
            string outPath = Path.Combine(imagesOutputDir, outFileName);

            // Save the image.
            shape.ImageData.Save(outPath);
            extractedCount++;
        }

        // Validation: ensure at least one image was extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // Optional: display result count (no interactive prompt required).
        Console.WriteLine($"Extracted {extractedCount} tile image(s) to \"{imagesOutputDir}\".");
    }
}
