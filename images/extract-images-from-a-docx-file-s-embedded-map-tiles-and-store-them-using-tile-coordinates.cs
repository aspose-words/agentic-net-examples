using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class ExtractMapTileImages
{
    public static void Main()
    {
        // Prepare folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "Input");
        string outputDir = Path.Combine(baseDir, "Output");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample tile images (PNG) with deterministic content
        CreateSampleTile(inputDir, 0, 0, Color.LightBlue);
        CreateSampleTile(inputDir, 0, 1, Color.LightGreen);
        CreateSampleTile(inputDir, 1, 0, Color.LightCoral);

        // Build a DOCX that contains the tiles as images
        string docPath = Path.Combine(inputDir, "MapTiles.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert each tile and store its coordinates in the shape's AlternativeText property
        InsertTile(builder, Path.Combine(inputDir, "tile_0_0.png"), 0, 0);
        InsertTile(builder, Path.Combine(inputDir, "tile_0_1.png"), 0, 1);
        InsertTile(builder, Path.Combine(inputDir, "tile_1_0.png"), 1, 0);

        // Save the document
        doc.Save(docPath);

        // Load the document and extract images using the stored coordinates
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int extractedCount = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Retrieve coordinates from AlternativeText (format: "x_y")
            string altText = shape.AlternativeText;
            if (string.IsNullOrEmpty(altText) || !altText.Contains("_"))
                continue;

            string[] parts = altText.Split('_');
            if (parts.Length != 2)
                continue;

            string x = parts[0];
            string y = parts[1];

            // Determine file extension based on image type
            string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
            string outFileName = $"tile_{x}_{y}{extension}";
            string outPath = Path.Combine(outputDir, outFileName);

            // Save the image
            shape.ImageData.Save(outPath);
            extractedCount++;
        }

        // Validation: ensure at least one image was extracted
        if (extractedCount == 0)
            throw new InvalidOperationException("No tile images were extracted from the document.");

        Console.WriteLine($"Extraction complete. {extractedCount} tile image(s) saved to '{outputDir}'.");
    }

    // Helper to create a simple PNG tile with a solid background color
    private static void CreateSampleTile(string folder, int x, int y, Color bgColor)
    {
        int width = 256;
        int height = 256;
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(bgColor);
            // Optionally draw the coordinates on the tile for visual verification
            // (Skipping text drawing to avoid font ambiguities per rules)
            string fileName = $"tile_{x}_{y}.png";
            string fullPath = Path.Combine(folder, fileName);
            bitmap.Save(fullPath, ImageFormat.Png);
        }
    }

    // Helper to insert an image and annotate it with its tile coordinates
    private static void InsertTile(DocumentBuilder builder, string imagePath, int x, int y)
    {
        Shape shape = builder.InsertImage(imagePath);
        // Store coordinates in AlternativeText for later extraction
        shape.AlternativeText = $"{x}_{y}";
        // Add a line break after each tile for readability
        builder.Writeln();
    }
}
