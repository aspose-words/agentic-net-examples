using System;
using System.IO;
using System.IO.Compression;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Prepare working directories
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "work");
        Directory.CreateDirectory(workDir);

        // 1. Create a sample TIFF image
        string sampleTiffPath = Path.Combine(workDir, "sample.tif");
        CreateSampleTiff(sampleTiffPath);

        // 2. Create an RTF document and embed the TIFF image
        string rtfPath = Path.Combine(workDir, "sample.rtf");
        CreateRtfWithImage(rtfPath, sampleTiffPath);

        // 3. Load the RTF document and extract images
        Document doc = new Document(rtfPath);
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        string[] compressedImagePaths = new string[shapeNodes.Count];
        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                // Save extracted image to a temporary file
                string extractedPath = Path.Combine(workDir, $"extracted_{imageIndex}.tif");
                using (MemoryStream imgStream = new MemoryStream())
                {
                    shape.ImageData.Save(imgStream);
                    imgStream.Position = 0;
                    File.WriteAllBytes(extractedPath, imgStream.ToArray());
                }

                // Apply lossless compression (re‑save as TIFF using default compression)
                string compressedPath = Path.Combine(workDir, $"compressed_{imageIndex}.tif");
                using (Bitmap bitmap = new Bitmap(extractedPath))
                {
                    // Re‑save the bitmap as TIFF (default compression is lossless)
                    bitmap.Save(compressedPath, ImageFormat.Tiff);
                }

                compressedImagePaths[imageIndex] = compressedPath;
                imageIndex++;
            }
        }

        if (imageIndex == 0)
            throw new Exception("No images were extracted from the RTF document.");

        // 4. Store compressed images in a ZIP archive
        string zipPath = Path.Combine(workDir, "images.zip");
        using (FileStream zipToOpen = new FileStream(zipPath, FileMode.Create))
        using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Update))
        {
            for (int i = 0; i < imageIndex; i++)
            {
                string filePath = compressedImagePaths[i];
                if (File.Exists(filePath))
                {
                    string entryName = Path.GetFileName(filePath);
                    archive.CreateEntryFromFile(filePath, entryName);
                }
            }
        }

        // Validation
        if (!File.Exists(zipPath))
            throw new Exception("ZIP archive was not created.");

        using (ZipArchive archive = ZipFile.OpenRead(zipPath))
        {
            if (archive.Entries.Count == 0)
                throw new Exception("ZIP archive contains no entries.");
        }

        // Clean up (optional)
        // Directory.Delete(workDir, true);
    }

    private static void CreateSampleTiff(string path)
    {
        int width = 200;
        int height = 200;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                using (Pen pen = new Pen(Aspose.Drawing.Color.Blue, 5))
                {
                    g.DrawEllipse(pen, 10, 10, width - 20, height - 20);
                }
                using (Brush brush = new SolidBrush(Aspose.Drawing.Color.Red))
                {
                    g.FillRectangle(brush, 50, 50, 100, 100);
                }
            }
            bitmap.Save(path, ImageFormat.Tiff);
        }
    }

    private static void CreateRtfWithImage(string rtfPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(rtfPath, SaveFormat.Rtf);
    }
}
