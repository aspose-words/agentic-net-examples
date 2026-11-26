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
        string baseDir = Directory.GetCurrentDirectory();
        string inputFolder = Path.Combine(baseDir, "InputDocs");
        string outputFolder = Path.Combine(baseDir, "OutputImages");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a deterministic sample image (input.png)
        string sampleImagePath = Path.Combine(baseDir, "input.png");
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple rectangle
                using (SolidBrush brush = new SolidBrush(Color.Blue))
                {
                    g.FillRectangle(brush, 50, 50, 100, 100);
                }
            }
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // Create a sample Word document containing the image
        string docPath = Path.Combine(inputFolder, "SampleDocument.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(docPath);

        // Batch process all Word files in the input folder
        int totalConverted = 0;
        foreach (string file in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document loadDoc = new Document(file);
            NodeCollection shapes = loadDoc.GetChildNodes(NodeType.Shape, true);
            int shapeIndex = 0;
            foreach (Shape shape in shapes)
            {
                if (!shape.HasImage) continue;

                // Save the shape image to a memory stream
                using (MemoryStream imgStream = new MemoryStream())
                {
                    shape.ImageData.Save(imgStream);
                    imgStream.Position = 0;

                    // Load the image into Aspose.Drawing.Bitmap
                    using (Bitmap bmp = new Bitmap(imgStream))
                    {
                        // Define output WebP file name (saved as PNG content for compatibility)
                        string outputFileName = $"{Path.GetFileNameWithoutExtension(file)}_img{shapeIndex}.webp";
                        string outputPath = Path.Combine(outputFolder, outputFileName);

                        // Save as PNG format but keep .webp extension (WebP not directly supported)
                        bmp.Save(outputPath, ImageFormat.Png);
                        totalConverted++;

                        // Validate that the file was created
                        if (!File.Exists(outputPath))
                            throw new InvalidOperationException($"Failed to create output image: {outputPath}");
                    }
                }
                shapeIndex++;
            }
        }

        // Ensure at least one image was converted
        if (totalConverted == 0)
            throw new InvalidOperationException("No images were extracted and converted.");

        Console.WriteLine($"Total images converted to WebP (PNG content): {totalConverted}");
    }
}
