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
        string inputFolder = Path.Combine(baseDir, "InputImages");
        string extractedFolder = Path.Combine(baseDir, "ExtractedImages");
        string outputFolder = Path.Combine(baseDir, "OutputImages");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(extractedFolder);
        Directory.CreateDirectory(outputFolder);

        // Step 1: Create sample TIFF images
        for (int i = 1; i <= 3; i++)
        {
            string tiffPath = Path.Combine(inputFolder, $"sample{i}.tiff");
            using (Bitmap bitmap = new Bitmap(200, 200))
            {
                using (Graphics g = Graphics.FromImage(bitmap))
                {
                    g.Clear(Aspose.Drawing.Color.White);
                    // Draw a simple rectangle with a different color for each image
                    Aspose.Drawing.Brush brush = new SolidBrush(Aspose.Drawing.Color.FromArgb(50 * i, 100, 150));
                    g.FillRectangle(brush, 20, 20, 160, 160);
                    brush.Dispose();
                }
                // Set DPI (example 300x300)
                bitmap.SetResolution(300f, 300f);
                bitmap.Save(tiffPath, ImageFormat.Tiff);
            }
        }

        // Step 2: Create a Word document and insert the TIFF images
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        foreach (string tiffFile in Directory.GetFiles(inputFolder, "*.tiff"))
        {
            builder.InsertImage(tiffFile);
            builder.Writeln(); // separate images
        }
        string docPath = Path.Combine(baseDir, "Sample.docx");
        doc.Save(docPath);

        // Step 3: Extract images from the document (they will be saved as TIFF)
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        int extractedIndex = 1;
        foreach (Shape shape in shapes)
        {
            if (shape.HasImage)
            {
                string extractedPath = Path.Combine(extractedFolder, $"extracted{extractedIndex}.tiff");
                shape.ImageData.Save(extractedPath);
                extractedIndex++;
            }
        }

        // Validate extraction
        if (Directory.GetFiles(extractedFolder, "*.tiff").Length == 0)
            throw new Exception("No images were extracted from the document.");

        // Step 4: Batch convert extracted TIFF images to lossless PNG while preserving DPI
        foreach (string tiffFile in Directory.GetFiles(extractedFolder, "*.tiff"))
        {
            using (Bitmap tiffBitmap = new Bitmap(tiffFile))
            {
                // Preserve DPI
                float dpiX = tiffBitmap.HorizontalResolution;
                float dpiY = tiffBitmap.VerticalResolution;

                // Create a new bitmap with the same size and DPI
                using (Bitmap pngBitmap = new Bitmap(tiffBitmap.Width, tiffBitmap.Height))
                {
                    pngBitmap.SetResolution(dpiX, dpiY);
                    using (Graphics g = Graphics.FromImage(pngBitmap))
                    {
                        g.Clear(Aspose.Drawing.Color.Transparent);
                        g.DrawImage(tiffBitmap, 0, 0, tiffBitmap.Width, tiffBitmap.Height);
                    }

                    string pngPath = Path.Combine(outputFolder,
                        Path.GetFileNameWithoutExtension(tiffFile) + ".png");
                    pngBitmap.Save(pngPath, ImageFormat.Png);
                }
            }
        }

        // Validate conversion
        if (Directory.GetFiles(outputFolder, "*.png").Length == 0)
            throw new Exception("No PNG files were created during conversion.");

        // Cleanup (optional)
        // Console.WriteLine("Batch conversion completed successfully.");
    }
}
