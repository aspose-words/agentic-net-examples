using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputFolder = Path.Combine(baseDir, "InputImages");
        string outputFolder = Path.Combine(baseDir, "OutputImages");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Step 1: Create sample BMP images
        List<string> sampleBmpPaths = new List<string>();
        for (int i = 1; i <= 2; i++)
        {
            string bmpPath = Path.Combine(baseDir, $"sample{i}.bmp");
            Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(200, 200);
            Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap);
            g.Clear(Aspose.Drawing.Color.White);
            // Draw a simple rectangle with a different color per image
            Aspose.Drawing.Color rectColor = i == 1 ? Aspose.Drawing.Color.Red : Aspose.Drawing.Color.Blue;
            using (Aspose.Drawing.Brush brush = new Aspose.Drawing.SolidBrush(rectColor))
            {
                g.FillRectangle(brush, 50, 50, 100, 100);
            }
            // Save BMP
            bitmap.Save(bmpPath, Aspose.Drawing.Imaging.ImageFormat.Bmp);
            // Cleanup
            g.Dispose();
            bitmap.Dispose();

            sampleBmpPaths.Add(bmpPath);
        }

        // Step 2: Create a Word document and insert the BMP images
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        foreach (string bmpPath in sampleBmpPaths)
        {
            builder.InsertImage(bmpPath);
            builder.Writeln(); // separate images
        }
        string docPath = Path.Combine(baseDir, "sample.docx");
        doc.Save(docPath);

        // Step 3: Load the document and extract embedded BMP images
        Document loadedDoc = new Document(docPath);
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;
        foreach (Shape shape in shapes)
        {
            if (shape.HasImage)
            {
                string extractedPath = Path.Combine(inputFolder, $"extracted-{extractedCount + 1}.bmp");
                shape.ImageData.Save(extractedPath);
                extractedCount++;
            }
        }

        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // Step 4: Convert each extracted BMP to lossless PNG (as a safe alternative to WebP)
        string[] bmpFiles = Directory.GetFiles(inputFolder, "*.bmp");
        if (bmpFiles.Length == 0)
            throw new InvalidOperationException("No BMP files found for conversion.");

        int convertedCount = 0;
        foreach (string bmpFile in bmpFiles)
        {
            using (Aspose.Drawing.Bitmap bmp = new Aspose.Drawing.Bitmap(bmpFile))
            {
                string pngFileName = Path.GetFileNameWithoutExtension(bmpFile) + ".png";
                string pngPath = Path.Combine(outputFolder, pngFileName);

                // Save as lossless PNG
                bmp.Save(pngPath, Aspose.Drawing.Imaging.ImageFormat.Png);

                // Log conversion details
                FileInfo originalInfo = new FileInfo(bmpFile);
                FileInfo pngInfo = new FileInfo(pngPath);
                Console.WriteLine($"Converted '{originalInfo.Name}' ({originalInfo.Length} bytes) to '{pngInfo.Name}' ({pngInfo.Length} bytes) losslessly.");

                convertedCount++;
            }
        }

        if (convertedCount == 0)
            throw new InvalidOperationException("No images were converted to PNG.");

        Console.WriteLine($"Batch conversion completed. {convertedCount} image(s) converted.");
    }
}
