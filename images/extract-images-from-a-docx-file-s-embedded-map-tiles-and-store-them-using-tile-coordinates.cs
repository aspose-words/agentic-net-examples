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
        // Prepare folders
        string baseDir = Directory.GetCurrentDirectory();
        string artifactsDir = Path.Combine(baseDir, "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create sample tile images
        CreateSampleTileImage(256, 256, Aspose.Drawing.Color.LightBlue, "tile_0_0.png", artifactsDir);
        CreateSampleTileImage(256, 256, Aspose.Drawing.Color.LightGreen, "tile_1_0.png", artifactsDir);

        // Build a DOCX containing the tiles with coordinate metadata in AlternativeText
        string docPath = Path.Combine(artifactsDir, "MapTiles.docx");
        BuildDocumentWithTiles(docPath, artifactsDir);

        // Load the document and extract images using tile coordinates
        Document doc = new Document(docPath);
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage) continue;

            // Expect AlternativeText in format "x_y"
            string altText = shape.AlternativeText;
            if (string.IsNullOrWhiteSpace(altText) || !altText.Contains("_"))
                continue;

            string[] parts = altText.Split('_');
            if (parts.Length != 2) continue;

            string x = parts[0];
            string y = parts[1];

            string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
            string fileName = $"tile_{x}_{y}{extension}";
            string outPath = Path.Combine(artifactsDir, fileName);

            shape.ImageData.Save(outPath);
            extractedCount++;
        }

        if (extractedCount == 0)
            throw new InvalidOperationException("No tile images were extracted from the document.");

        Console.WriteLine($"Extraction complete. {extractedCount} image(s) saved to '{artifactsDir}'.");
    }

    private static void CreateSampleTileImage(int width, int height, Aspose.Drawing.Color fillColor, string fileName, string folder)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(fillColor);
            string path = Path.Combine(folder, fileName);
            bitmap.Save(path, ImageFormat.Png);
        }
    }

    private static void BuildDocumentWithTiles(string docPath, string imagesFolder)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First tile (0,0)
        InsertTile(builder, Path.Combine(imagesFolder, "tile_0_0.png"), "0_0");

        // Second tile (1,0)
        InsertTile(builder, Path.Combine(imagesFolder, "tile_1_0.png"), "1_0");

        doc.Save(docPath, SaveFormat.Docx);
    }

    private static void InsertTile(DocumentBuilder builder, string imagePath, string coordinates)
    {
        Shape shape = builder.InsertImage(imagePath);
        shape.AlternativeText = coordinates; // Store tile coordinates
        shape.WrapType = WrapType.Inline;
        // Add a line break after each tile for readability
        builder.Writeln();
    }
}
