using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Folder for all generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a deterministic sample image that will be used as a map tile.
        // -----------------------------------------------------------------
        string tileImagePath = Path.Combine(outputDir, "tile.png");
        CreateSampleTileImage(tileImagePath, 100, 100);

        // -----------------------------------------------------------------
        // 2. Build a DOCX document containing several tiles placed at different coordinates.
        // -----------------------------------------------------------------
        string docPath = Path.Combine(outputDir, "MapTiles.docx");
        BuildDocumentWithTiles(docPath, tileImagePath);

        // -----------------------------------------------------------------
        // 3. Load the document and extract each tile image, naming the file by its tile coordinates.
        // -----------------------------------------------------------------
        Document doc = new Document(docPath);
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int extractedCount = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Determine tile coordinates from the shape's position.
            // The shape's Left/Top are in points; we placed each tile with a width/height of 100 points.
            int tileX = (int)Math.Round(shape.Left / shape.Width);
            int tileY = (int)Math.Round(shape.Top / shape.Height);

            // Build a deterministic file name using the coordinates and the image's original extension.
            string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
            string imageFileName = $"tile_{tileX}_{tileY}{extension}";
            string imageFullPath = Path.Combine(outputDir, imageFileName);

            // Save the image.
            shape.ImageData.Save(imageFullPath);
            extractedCount++;
        }

        // Validation: at least one image must have been extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the document.");
    }

    // Creates a simple PNG image with a solid background.
    private static void CreateSampleTileImage(string filePath, int width, int height)
    {
        // Ensure any previous file is removed.
        if (File.Exists(filePath))
            File.Delete(filePath);

        // Create bitmap and draw a deterministic background.
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.LightGray);
        // Save and dispose.
        bitmap.Save(filePath);
        graphics.Dispose();
        bitmap.Dispose();
    }

    // Builds a document containing a 3x3 grid of tiles.
    private static void BuildDocumentWithTiles(string docPath, string tileImagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define tile size (points). 1 point = 1/72 inch.
        const double tileSize = 100.0;

        // Insert tiles at coordinates (x, y) where x and y range from 0 to 2.
        for (int y = 0; y < 3; y++)
        {
            for (int x = 0; x < 3; x++)
            {
                // Insert the image and obtain the created shape.
                Shape tileShape = builder.InsertImage(tileImagePath);
                tileShape.WrapType = WrapType.None;
                tileShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
                tileShape.RelativeVerticalPosition = RelativeVerticalPosition.Page;

                // Set explicit size and position to simulate map tile coordinates.
                tileShape.Width = tileSize;
                tileShape.Height = tileSize;
                tileShape.Left = x * tileSize;
                tileShape.Top = y * tileSize;

                // No need to append the shape manually; InsertImage already adds it to the document.
            }
        }

        // Save the document.
        doc.Save(docPath);
    }
}
