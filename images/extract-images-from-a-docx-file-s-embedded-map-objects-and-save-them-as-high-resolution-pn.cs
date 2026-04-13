using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);
        string inputDocPath = Path.Combine(workDir, "input.docx");

        // ---------- Create a sample image (simulating a map) ----------
        string mapImagePath = Path.Combine(workDir, "map.png");
        int mapWidth = 800;
        int mapHeight = 600;
        using (Bitmap bitmap = new Bitmap(mapWidth, mapHeight))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            // Simple deterministic drawing to represent a map
            g.Clear(Aspose.Drawing.Color.White);
            g.DrawRectangle(new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 5), 50, 50, mapWidth - 100, mapHeight - 100);
            g.DrawString("Sample Map", new Aspose.Drawing.Font("Arial", 48), new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.DarkGreen), new Aspose.Drawing.PointF(200, 250));
            bitmap.Save(mapImagePath);
        }

        // ---------- Create a DOCX and embed the map image ----------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document with embedded map image:");
        builder.InsertImage(mapImagePath);
        doc.Save(inputDocPath);

        // ---------- Load the document with options to convert metafiles to PNG ----------
        LoadOptions loadOptions = new LoadOptions { ConvertMetafilesToPng = true };
        Document loadedDoc = new Document(inputDocPath, loadOptions);

        // ---------- Extract all images from shapes and save as high‑resolution PNG ----------
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Save image data to a memory stream
            using (MemoryStream imgStream = new MemoryStream())
            {
                shape.ImageData.Save(imgStream);
                imgStream.Position = 0;

                // Load into Aspose.Drawing.Bitmap to ensure PNG output
                using (Bitmap bmp = new Bitmap(imgStream))
                {
                    string outFile = Path.Combine(workDir, $"extracted_image_{imageIndex}.png");
                    bmp.Save(outFile);
                    Console.WriteLine($"Saved image {imageIndex} to: {outFile}");
                }
            }
            imageIndex++;
        }

        // ---------- Validation ----------
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        Console.WriteLine("Image extraction completed successfully.");
    }
}
