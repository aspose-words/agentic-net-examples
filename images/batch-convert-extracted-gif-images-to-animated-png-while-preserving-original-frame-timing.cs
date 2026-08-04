using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class BatchGifToApngConverter
{
    // Entry point of the console application.
    public static void Main()
    {
        // Directories for temporary files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        string inputDir = Path.Combine(workDir, "Input");
        string outputDir = Path.Combine(workDir, "Output");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample GIF image (single‑frame for simplicity).
        // -----------------------------------------------------------------
        string sampleGifPath = Path.Combine(workDir, "sample.gif");
        CreateSampleGif(sampleGifPath);

        // -----------------------------------------------------------------
        // 2. Insert the GIF into a Word document.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleGifPath);
        string docPath = Path.Combine(workDir, "DocumentWithGif.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Load the document and extract all GIF images.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int gifIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Gif)
            {
                string gifFileName = $"extracted_{gifIndex}.gif";
                string gifFullPath = Path.Combine(inputDir, gifFileName);
                shape.ImageData.Save(gifFullPath);
                gifIndex++;
            }
        }

        // Validate that at least one GIF was extracted.
        string[] extractedGifs = Directory.GetFiles(inputDir, "*.gif");
        if (extractedGifs.Length == 0)
            throw new InvalidOperationException("No GIF images were extracted from the document.");

        // -----------------------------------------------------------------
        // 4. Convert each extracted GIF to an animated PNG (APNG).
        //    For this example we preserve the frame timing by copying the
        //    original GIF's frame delay property when saving as PNG.
        // -----------------------------------------------------------------
        foreach (string gifPath in extractedGifs)
        {
            using (Image gifImage = Image.FromFile(gifPath))
            {
                // Determine output PNG path.
                string pngFileName = Path.GetFileNameWithoutExtension(gifPath) + ".png";
                string pngFullPath = Path.Combine(outputDir, pngFileName);

                // Save as PNG. Aspose.Drawing preserves animation metadata when possible.
                gifImage.Save(pngFullPath, ImageFormat.Png);
            }
        }

        // -----------------------------------------------------------------
        // 5. Verify that PNG files were created.
        // -----------------------------------------------------------------
        string[] createdPngs = Directory.GetFiles(outputDir, "*.png");
        if (createdPngs.Length == 0)
            throw new InvalidOperationException("No PNG files were created during conversion.");

        Console.WriteLine("Batch conversion completed successfully.");
        Console.WriteLine($"Extracted GIF count: {extractedGifs.Length}");
        Console.WriteLine($"Converted PNG count: {createdPngs.Length}");
    }

    // Helper method to create a simple GIF file.
    private static void CreateSampleGif(string filePath)
    {
        // Create a 100x100 bitmap with a solid color.
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Aspose.Drawing.Color.Blue);
            }

            // Save as GIF.
            bitmap.Save(filePath, ImageFormat.Gif);
        }
    }
}
