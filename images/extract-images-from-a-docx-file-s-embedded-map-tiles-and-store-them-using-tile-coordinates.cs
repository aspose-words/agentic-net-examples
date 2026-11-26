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
        // Prepare a deterministic output folder.
        string outputFolder = "output";
        Directory.CreateDirectory(outputFolder);

        // Sample tile coordinates.
        var tileCoordinates = new List<(int X, int Y)>
        {
            (0, 0), (0, 1), (1, 0), (1, 1)
        };

        // Create sample tile images and insert them into a DOCX.
        string docPath = Path.Combine(outputFolder, "map_tiles.docx");
        CreateDocumentWithTiles(docPath, tileCoordinates, outputFolder);

        // Load the document and extract images using tile coordinates.
        ExtractTiles(docPath, outputFolder);
    }

    private static void CreateDocumentWithTiles(string docPath, List<(int X, int Y)> tiles, string folder)
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        foreach (var (x, y) in tiles)
        {
            // Create a deterministic sample image for the tile.
            string imageFileName = $"tile_{x}_{y}.png";
            string imagePath = Path.Combine(folder, imageFileName);
            CreateSampleTileImage(imagePath, x, y);

            // Insert the image into the document.
            builder.InsertImage(imagePath);

            // The inserted shape is the last child of the current paragraph.
            Shape shape = builder.CurrentParagraph.LastChild as Shape;
            if (shape != null && shape.HasImage)
            {
                // Store tile coordinates in the shape's AlternativeText property.
                shape.AlternativeText = $"{x}_{y}";
            }

            // Add a line break between tiles for readability.
            builder.Writeln();
        }

        // Save the document containing the map tiles.
        doc.Save(docPath);
    }

    private static void CreateSampleTileImage(string path, int x, int y)
    {
        const int width = 100;
        const int height = 100;

        // Create a bitmap and draw deterministic content.
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                // Fill background with white.
                graphics.Clear(Aspose.Drawing.Color.White);

                // Draw a simple rectangle with a color based on coordinates.
                Aspose.Drawing.Color fillColor = Aspose.Drawing.Color.FromArgb(255, (x * 80) % 256, (y * 80) % 256);
                using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(fillColor))
                {
                    graphics.FillRectangle(brush, 10, 10, width - 20, height - 20);
                }

                // Draw the coordinates as text using Aspose.Drawing.Font.
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 12))
                using (Aspose.Drawing.SolidBrush textBrush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black))
                {
                    graphics.DrawString($"{x},{y}", font, textBrush, new Aspose.Drawing.PointF(15, 40));
                }
            }

            // Save the bitmap as PNG.
            bitmap.Save(path, Aspose.Drawing.Imaging.ImageFormat.Png);
        }
    }

    private static void ExtractTiles(string docPath, string outputFolder)
    {
        // Load the document containing the map tiles.
        Document doc = new Document(docPath);

        // Find all Shape nodes in the document.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            // Retrieve tile coordinates from AlternativeText.
            string altText = shape.AlternativeText;
            if (string.IsNullOrEmpty(altText))
                continue;

            string[] parts = altText.Split('_');
            if (parts.Length != 2)
                continue;

            if (!int.TryParse(parts[0], out int x) || !int.TryParse(parts[1], out int y))
                continue;

            // Save the extracted image using tile coordinates.
            string outputImagePath = Path.Combine(outputFolder, $"extracted_tile_{x}_{y}.png");
            shape.ImageData.Save(outputImagePath);
            extractedCount++;
        }

        // Validate that at least one image was extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No tile images were extracted from the document.");
    }
}
