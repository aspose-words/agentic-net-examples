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
        // Prepare deterministic folders
        const string inputFolder = "InputImages";
        const string outputFolder = "ExtractedTiles";
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Step 1: Create sample tile images (2x2 grid)
        for (int x = 0; x < 2; x++)
        {
            for (int y = 0; y < 2; y++)
            {
                string fileName = $"tile_{x}_{y}.png";
                string filePath = Path.Combine(inputFolder, fileName);

                // Create a simple bitmap
                using (Bitmap bitmap = new Bitmap(100, 100))
                {
                    using (Graphics g = Graphics.FromImage(bitmap))
                    {
                        // Fill background with a distinct color per tile
                        Color bgColor = (x + y) % 2 == 0 ? Color.LightBlue : Color.LightGreen;
                        g.Clear(bgColor);
                    }

                    // Save the bitmap to a deterministic file
                    bitmap.Save(filePath, ImageFormat.Png);
                }
            }
        }

        // Step 2: Build a DOCX that contains the tile images
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int x = 0; x < 2; x++)
        {
            for (int y = 0; y < 2; y++)
            {
                string imagePath = Path.Combine(inputFolder, $"tile_{x}_{y}.png");

                // Insert the image and obtain the created Shape
                Shape shape = builder.InsertImage(imagePath);
                // Store tile coordinates in AlternativeText for later extraction
                shape.AlternativeText = $"{x}_{y}";

                // Add a line break between images for readability
                builder.Writeln();
            }
        }

        // Save the sample document
        const string docPath = "MapTiles.docx";
        doc.Save(docPath);

        // Step 3: Load the document and extract images using tile coordinates
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int extractedCount = 0;
        int genericIndex = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            // Determine output file name based on stored coordinates
            string altText = shape.AlternativeText;
            string outputFileName;

            if (!string.IsNullOrWhiteSpace(altText))
            {
                // Expected format: "x_y"
                outputFileName = $"tile_{altText}.png";
            }
            else
            {
                // Fallback name if no coordinates are stored
                outputFileName = $"extracted_{genericIndex++}.png";
            }

            string outputPath = Path.Combine(outputFolder, outputFileName);
            shape.ImageData.Save(outputPath);
            extractedCount++;
        }

        // Validation: ensure at least one image was extracted
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        Console.WriteLine($"Extraction complete. {extractedCount} image(s) saved to '{outputFolder}'.");
    }
}
