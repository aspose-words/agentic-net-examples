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
        // Prepare output directory.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -------------------------------------------------------------
        // 1. Create a deterministic sample PNG image using Aspose.Drawing.
        // -------------------------------------------------------------
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                g.FillRectangle(new SolidBrush(Color.Blue), 20, 20, 160, 160);
            }
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // -------------------------------------------------------------
        // 2. Build an RTF document and embed the sample image.
        // -------------------------------------------------------------
        Document rtfDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(rtfDoc);
        builder.InsertImage(sampleImagePath);
        string rtfPath = Path.Combine(artifactsDir, "sample.rtf");
        rtfDoc.Save(rtfPath, SaveFormat.Rtf);

        // -------------------------------------------------------------
        // 3. Load the RTF document and extract all images.
        // -------------------------------------------------------------
        Document loadedDoc = new Document(rtfPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        string zipPath = Path.Combine(artifactsDir, "ImagesArchive.zip");
        using (FileStream zipStream = new FileStream(zipPath, FileMode.Create))
        using (ZipArchive archive = new ZipArchive(zipStream, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage)
                    continue;

                // -------------------------------------------------
                // 4. Save the original image to a memory stream.
                // -------------------------------------------------
                using (MemoryStream imgStream = new MemoryStream())
                {
                    shape.ImageData.Save(imgStream);
                    imgStream.Position = 0; // Reset before reading.

                    // -------------------------------------------------
                    // 5. Convert the image to TIFF with lossless LZW compression.
                    // -------------------------------------------------
                    string tiffFileName = $"image_{imageIndex}.tif";
                    ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
                    {
                        TiffCompression = TiffCompression.Lzw // LZW is lossless.
                    };

                    using (MemoryStream tiffStream = new MemoryStream())
                    {
                        // Render the image as a single‑page document and save as TIFF.
                        Document tempDoc = new Document();
                        DocumentBuilder tempBuilder = new DocumentBuilder(tempDoc);
                        tempBuilder.InsertImage(imgStream);
                        tempDoc.Save(tiffStream, tiffOptions);
                        tiffStream.Position = 0;

                        // -------------------------------------------------
                        // 6. Add the TIFF image to the ZIP archive.
                        // -------------------------------------------------
                        ZipArchiveEntry entry = archive.CreateEntry(tiffFileName, System.IO.Compression.CompressionLevel.Optimal);
                        using (Stream entryStream = entry.Open())
                        {
                            tiffStream.CopyTo(entryStream);
                        }
                    }

                    imageIndex++;
                }
            }
        }

        // -------------------------------------------------------------
        // 7. Validation.
        // -------------------------------------------------------------
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the RTF document.");

        if (!File.Exists(zipPath))
            throw new FileNotFoundException("The archive file was not created.", zipPath);

        Console.WriteLine($"RTF document saved to: {rtfPath}");
        Console.WriteLine($"Archive with TIFF images saved to: {zipPath}");
    }
}
