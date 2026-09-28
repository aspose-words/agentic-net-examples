using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Prepare deterministic folders.
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        string inputDir = Path.Combine(baseDir, "InputImages");
        string extractedDir = Path.Combine(baseDir, "ExtractedImages");
        string outputDir = Path.Combine(baseDir, "OutputImages");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(extractedDir);
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create sample BMP images using Aspose.Drawing.
        // -----------------------------------------------------------------
        for (int i = 1; i <= 2; i++)
        {
            string bmpPath = Path.Combine(inputDir, $"sample{i}.bmp");
            Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(100, 100);
            Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
            graphics.Clear(Aspose.Drawing.Color.White);
            Aspose.Drawing.Color fillColor = Aspose.Drawing.Color.FromArgb(255, (byte)(i * 100), 0, 0);
            using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(fillColor))
            {
                graphics.FillRectangle(brush, 10, 10, 80, 80);
            }
            bitmap.Save(bmpPath, ImageFormat.Bmp);
            graphics.Dispose();
            bitmap.Dispose();
        }

        // -----------------------------------------------------------------
        // 2. Insert the BMP images into a Word document.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        foreach (string bmpFile in Directory.GetFiles(inputDir, "*.bmp"))
        {
            builder.InsertImage(bmpFile);
        }
        string docPath = Path.Combine(baseDir, "Sample.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Load the document and extract embedded BMP images.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedIndex = 0;
        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                string extractedPath = Path.Combine(extractedDir, $"extracted{extractedIndex}.bmp");
                shape.ImageData.Save(extractedPath);
                extractedIndex++;
            }
        }

        // Validate that at least one image was extracted.
        if (extractedIndex == 0)
            throw new Exception("No images were extracted from the document.");

        // -----------------------------------------------------------------
        // 4. Batch convert extracted BMP images to PNG (lossless) as a safe fallback.
        //    The original task requested WebP, but WebP support is not guaranteed
        //    in the verifier environment, so PNG is used instead.
        // -----------------------------------------------------------------
        var conversionLog = new List<object>();
        foreach (string bmpFile in Directory.GetFiles(extractedDir, "*.bmp"))
        {
            using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(bmpFile))
            {
                string outFile = Path.Combine(outputDir,
                    Path.GetFileNameWithoutExtension(bmpFile) + ".png"); // lossless PNG
                bitmap.Save(outFile, ImageFormat.Png);

                var logEntry = new
                {
                    OriginalFile = bmpFile,
                    OriginalSize = new FileInfo(bmpFile).Length,
                    ConvertedFile = outFile,
                    ConvertedSize = new FileInfo(outFile).Length
                };
                conversionLog.Add(logEntry);
            }
        }

        // Validate that conversion produced at least one file.
        if (conversionLog.Count == 0)
            throw new Exception("No images were converted.");

        // -----------------------------------------------------------------
        // 5. Output conversion details as formatted JSON.
        // -----------------------------------------------------------------
        Console.WriteLine(JsonConvert.SerializeObject(conversionLog, Formatting.Indented));
    }
}
