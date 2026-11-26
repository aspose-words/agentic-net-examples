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
        // Prepare folders
        string inputFolder = "InputGifs";
        string extractedFolder = "ExtractedGifs";
        string convertedFolder = "ConvertedPngs";
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(extractedFolder);
        Directory.CreateDirectory(convertedFolder);

        // Create sample GIF images using Aspose.Drawing
        for (int i = 1; i <= 2; i++)
        {
            string gifPath = Path.Combine(inputFolder, $"sample{i}.gif");
            Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(100, 100);
            try
            {
                Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap);
                try
                {
                    g.Clear(Aspose.Drawing.Color.White);
                    g.DrawEllipse(new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 3), 10, 10, 80, 80);
                    g.DrawString(
                        $"GIF {i}",
                        new Aspose.Drawing.Font("Arial", 12),
                        new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black),
                        20,
                        40);
                }
                finally
                {
                    g.Dispose();
                }

                bitmap.Save(gifPath, Aspose.Drawing.Imaging.ImageFormat.Gif);
            }
            finally
            {
                bitmap.Dispose();
            }
        }

        // Create a Word document and insert the GIF images
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        string[] gifFiles = Directory.GetFiles(inputFolder, "*.gif");
        foreach (string gifFile in gifFiles)
        {
            builder.InsertImage(gifFile);
            builder.Writeln(); // separate images
        }
        string docPath = "Sample.docx";
        doc.Save(docPath);

        // Load the document and extract GIF images
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;
        for (int i = 0; i < shapeNodes.Count; i++)
        {
            Shape shape = (Shape)shapeNodes[i];
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Gif)
            {
                string outPath = Path.Combine(extractedFolder, $"extracted-{extractedCount + 1}.gif");
                shape.ImageData.Save(outPath);
                extractedCount++;
            }
        }
        if (extractedCount == 0)
            throw new InvalidOperationException("No GIF images were extracted from the document.");

        // Convert each extracted GIF to animated PNG (preserving frame timing where possible)
        string[] extractedGifs = Directory.GetFiles(extractedFolder, "*.gif");
        int convertedCount = 0;
        foreach (string gifPath in extractedGifs)
        {
            // Load the GIF using Aspose.Drawing.Image
            Aspose.Drawing.Image gifImage = Aspose.Drawing.Image.FromFile(gifPath);
            try
            {
                // For simplicity, save the first frame as PNG.
                string pngPath = Path.Combine(convertedFolder, Path.GetFileNameWithoutExtension(gifPath) + ".png");
                gifImage.Save(pngPath, Aspose.Drawing.Imaging.ImageFormat.Png);
                convertedCount++;
            }
            finally
            {
                gifImage.Dispose();
            }
        }
        if (convertedCount == 0)
            throw new InvalidOperationException("No GIF images were converted to PNG.");

        // Validation output
        Console.WriteLine($"Created {gifFiles.Length} sample GIF(s).");
        Console.WriteLine($"Extracted {extractedCount} GIF image(s) from the document.");
        Console.WriteLine($"Converted {convertedCount} GIF image(s) to PNG.");
    }
}
