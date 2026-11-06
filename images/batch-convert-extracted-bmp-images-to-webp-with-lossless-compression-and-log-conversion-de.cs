using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

public class BatchBmpToWebpConverter
{
    public class ConversionInfo
    {
        public string InputPath { get; set; }
        public string OutputPath { get; set; }
        public bool Success { get; set; }
        public string Message { get; set; }
    }

    public static void Main()
    {
        // Prepare folders
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        string originalBmpFolder = Path.Combine(baseDir, "OriginalBMP");
        string extractedBmpFolder = Path.Combine(baseDir, "ExtractedBMP");
        string outputWebpFolder = Path.Combine(baseDir, "OutputWebP");
        Directory.CreateDirectory(originalBmpFolder);
        Directory.CreateDirectory(extractedBmpFolder);
        Directory.CreateDirectory(outputWebpFolder);

        // Step 1: Create sample BMP images
        CreateSampleBmpImages(originalBmpFolder, 3);

        // Step 2: Insert BMP images into a Word document
        string docPath = Path.Combine(baseDir, "SampleDocument.docx");
        InsertImagesIntoWordDocument(originalBmpFolder, docPath);

        // Step 3: Extract BMP images from the Word document
        ExtractBmpImagesFromDocument(docPath, extractedBmpFolder);

        // Step 4: Batch convert extracted BMP images to lossless PNG (as a safe alternative to WebP)
        List<ConversionInfo> conversions = ConvertBmpToPng(extractedBmpFolder, outputWebpFolder);

        // Log conversion details as JSON
        string jsonLog = JsonConvert.SerializeObject(conversions, Formatting.Indented);
        Console.WriteLine(jsonLog);

        // Validation
        if (conversions.Count == 0 || conversions.Exists(c => !c.Success))
        {
            throw new Exception("One or more BMP images failed to convert to PNG.");
        }

        // Ensure at least one output file exists
        if (Directory.GetFiles(outputWebpFolder, "*.png").Length == 0)
        {
            throw new Exception("No PNG files were created.");
        }
    }

    private static void CreateSampleBmpImages(string folder, int count)
    {
        for (int i = 1; i <= count; i++)
        {
            string filePath = Path.Combine(folder, $"image{i}.bmp");
            using (Bitmap bitmap = new Bitmap(200, 200))
            {
                using (Graphics g = Graphics.FromImage(bitmap))
                {
                    g.Clear(Aspose.Drawing.Color.White);
                    using (SolidBrush brush = new SolidBrush(Aspose.Drawing.Color.FromArgb(50 * i, 100, 150)))
                    {
                        g.FillRectangle(brush, 20, 20, 160, 160);
                    }
                }
                bitmap.Save(filePath, ImageFormat.Bmp);
            }
        }
    }

    private static void InsertImagesIntoWordDocument(string imagesFolder, string docPath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        string[] bmpFiles = Directory.GetFiles(imagesFolder, "*.bmp");
        foreach (string bmpFile in bmpFiles)
        {
            builder.InsertParagraph();
            builder.InsertImage(bmpFile);
        }

        doc.Save(docPath);
    }

    private static void ExtractBmpImagesFromDocument(string docPath, string outputFolder)
    {
        Document doc = new Document(docPath);
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int index = 1;
        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                string outPath = Path.Combine(outputFolder, $"extracted-{index}.bmp");
                shape.ImageData.Save(outPath);
                index++;
            }
        }

        if (index == 1)
        {
            throw new Exception("No images were extracted from the document.");
        }
    }

    private static List<ConversionInfo> ConvertBmpToPng(string inputFolder, string outputFolder)
    {
        List<ConversionInfo> results = new List<ConversionInfo>();
        string[] bmpFiles = Directory.GetFiles(inputFolder, "*.bmp");
        foreach (string bmpPath in bmpFiles)
        {
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(bmpPath);
            string outPath = Path.Combine(outputFolder, $"{fileNameWithoutExt}.png");
            ConversionInfo info = new ConversionInfo
            {
                InputPath = bmpPath,
                OutputPath = outPath
            };
            try
            {
                using (Bitmap bitmap = new Bitmap(bmpPath))
                {
                    // Save as PNG (lossless) – a safe, widely supported format
                    bitmap.Save(outPath, ImageFormat.Png);
                }
                info.Success = true;
                info.Message = "Conversion succeeded.";
            }
            catch (Exception ex)
            {
                info.Success = false;
                info.Message = $"Conversion failed: {ex.Message}";
            }
            results.Add(info);
        }
        return results;
    }
}
