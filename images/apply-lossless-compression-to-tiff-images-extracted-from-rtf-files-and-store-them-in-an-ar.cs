using System;
using System.IO;
using System.IO.Compression;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Directory for all generated files.
        string artifactsDir = "Artifacts";
        Directory.CreateDirectory(artifactsDir);

        // -------------------------------------------------
        // 1. Create a deterministic sample image (PNG).
        // -------------------------------------------------
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Aspose.Drawing.Color.LightBlue);
            }
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // -------------------------------------------------
        // 2. Insert the image into a document and save as RTF.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        string rtfPath = Path.Combine(artifactsDir, "sample.rtf");
        doc.Save(rtfPath, SaveFormat.Rtf);

        // -------------------------------------------------
        // 3. Load the RTF document and extract all images.
        // -------------------------------------------------
        Document rtfDoc = new Document(rtfPath);
        NodeCollection shapeNodes = rtfDoc.GetChildNodes(NodeType.Shape, true);

        // -------------------------------------------------
        // 4. Convert each extracted image to TIFF with lossless (LZW) compression
        //    and store the TIFF files in a ZIP archive.
        // -------------------------------------------------
        string zipPath = Path.Combine(artifactsDir, "ImagesArchive.zip");
        int extractedCount = 0;

        using (FileStream zipStream = new FileStream(zipPath, FileMode.Create))
        using (ZipArchive archive = new ZipArchive(zipStream, ZipArchiveMode.Update))
        {
            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage)
                    continue;

                // Save the shape's image to a memory stream.
                using (MemoryStream imageStream = new MemoryStream())
                {
                    shape.ImageData.Save(imageStream);
                    imageStream.Position = 0; // Reset before reading.

                    // Load the image into a bitmap.
                    using (Bitmap bitmap = new Bitmap(imageStream))
                    {
                        // Save the bitmap as TIFF. Default LZW compression is lossless.
                        string tiffFileName = $"image_{extractedCount}.tiff";
                        string tiffFullPath = Path.Combine(artifactsDir, tiffFileName);
                        bitmap.Save(tiffFullPath, ImageFormat.Tiff);

                        // Add the TIFF file to the ZIP archive.
                        archive.CreateEntryFromFile(tiffFullPath, tiffFileName);
                        extractedCount++;
                    }
                }
            }
        }

        // -------------------------------------------------
        // 5. Validation – ensure at least one image was archived.
        // -------------------------------------------------
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted and archived.");
    }
}
