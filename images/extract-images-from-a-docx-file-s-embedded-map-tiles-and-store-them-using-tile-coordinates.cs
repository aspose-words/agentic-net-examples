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
        // Folder for temporary files
        string workFolder = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workFolder);

        // Create sample tile images (2x2 grid)
        int tileCountX = 2;
        int tileCountY = 2;
        int tileSize = 100; // pixels

        for (int x = 0; x < tileCountX; x++)
        {
            for (int y = 0; y < tileCountY; y++)
            {
                string tilePath = Path.Combine(workFolder, $"tile_{x}_{y}.png");
                using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(tileSize, tileSize))
                {
                    using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
                    {
                        g.Clear(Aspose.Drawing.Color.White);
                        // Draw a simple rectangle with the coordinates
                        g.DrawRectangle(Aspose.Drawing.Pens.Black, 10, 10, tileSize - 20, tileSize - 20);
                        g.DrawString($"({x},{y})",
                            new Aspose.Drawing.Font("Arial", 12),
                            Aspose.Drawing.Brushes.Black,
                            new Aspose.Drawing.PointF(20, 40));
                    }
                    bitmap.Save(tilePath, Aspose.Drawing.Imaging.ImageFormat.Png);
                }
            }
        }

        // Create a DOCX and insert the tiles, storing coordinates in AlternativeText
        string docPath = Path.Combine(workFolder, "MapTiles.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int x = 0; x < tileCountX; x++)
        {
            for (int y = 0; y < tileCountY; y++)
            {
                string tilePath = Path.Combine(workFolder, $"tile_{x}_{y}.png");
                Shape shape = builder.InsertImage(tilePath);
                shape.AlternativeText = $"tile_{x}_{y}";
                // Add a line break after each image for readability
                builder.Writeln();
            }
        }

        doc.Save(docPath);

        // Load the document and extract images using tile coordinates from AlternativeText
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            string altText = shape.AlternativeText ?? string.Empty;
            if (!altText.StartsWith("tile_"))
                continue;

            string[] parts = altText.Split('_');
            if (parts.Length != 3)
                continue;

            string xPart = parts[1];
            string yPart = parts[2];

            string outputFileName = Path.Combine(workFolder, $"extracted_tile_{xPart}_{yPart}.png");
            shape.ImageData.Save(outputFileName);
            extractedCount++;
        }

        if (extractedCount == 0)
            throw new InvalidOperationException("No tile images were extracted from the document.");

        // Optional: clean up sample tile images (comment out if you want to keep them)
        //foreach (var file in Directory.GetFiles(workFolder, "tile_*.png"))
        //    File.Delete(file);
    }
}
