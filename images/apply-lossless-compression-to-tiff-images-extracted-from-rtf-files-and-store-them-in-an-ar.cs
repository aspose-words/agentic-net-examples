using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Directories for artifacts
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        string imagesDir = Path.Combine(artifactsDir, "Images");
        Directory.CreateDirectory(artifactsDir);
        Directory.CreateDirectory(imagesDir);

        // 1. Create a sample PNG image using Aspose.Drawing
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        CreateSamplePng(sampleImagePath);

        // 2. Build an RTF document that contains the sample image
        string rtfPath = Path.Combine(artifactsDir, "sample.rtf");
        CreateRtfWithImage(sampleImagePath, rtfPath);

        // 3. Load the RTF document and extract images, converting each to a losslessly compressed TIFF
        string tiffOutputDir = Path.Combine(artifactsDir, "TiffImages");
        Directory.CreateDirectory(tiffOutputDir);
        ExtractImagesToTiff(rtfPath, tiffOutputDir);

        // 4. Verify that at least one TIFF file was created
        var tiffFiles = Directory.GetFiles(tiffOutputDir, "*.tiff");
        if (tiffFiles.Length == 0)
            throw new InvalidOperationException("No TIFF images were generated.");

        // 5. Archive the TIFF images into a ZIP file
        string zipPath = Path.Combine(artifactsDir, "TiffImages.zip");
        if (File.Exists(zipPath))
            File.Delete(zipPath);
        ZipFile.CreateFromDirectory(tiffOutputDir, zipPath);

        // Verify archive creation
        if (!File.Exists(zipPath))
            throw new InvalidOperationException("Failed to create the ZIP archive.");
    }

    private static void CreateSamplePng(string filePath)
    {
        // Create a 100x100 white bitmap and save as PNG
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(100, 100))
        using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
        {
            graphics.Clear(Aspose.Drawing.Color.White);
            bitmap.Save(filePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        }
    }

    private static void CreateRtfWithImage(string imagePath, string rtfPath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(rtfPath, SaveFormat.Rtf);
    }

    private static void ExtractImagesToTiff(string rtfPath, string outputDir)
    {
        Document doc = new Document(rtfPath);
        var shapes = doc.GetChildNodes(NodeType.Shape, true)
                        .OfType<Shape>()
                        .Where(s => s.HasImage)
                        .ToList();

        int index = 0;
        foreach (Shape shape in shapes)
        {
            // Obtain the image bytes from the shape
            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0; // Reset before reuse

                // Create a temporary document that contains only this image
                Document tempDoc = new Document();
                DocumentBuilder tempBuilder = new DocumentBuilder(tempDoc);
                tempBuilder.InsertImage(imageStream);

                // Configure TIFF save options with lossless LZW compression
                ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
                {
                    TiffCompression = TiffCompression.Lzw
                };

                string tiffPath = Path.Combine(outputDir, $"image_{index}.tiff");
                tempDoc.Save(tiffPath, options);
                index++;
            }
        }
    }
}
