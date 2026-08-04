using System;
using System.IO;
using System.IO.Compression;
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
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a sample image using Aspose.Drawing (deterministic PNG)
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        using (Bitmap bitmap = new Bitmap(200, 100))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            g.Clear(Aspose.Drawing.Color.White);
            // Draw a simple rectangle
            g.FillRectangle(new SolidBrush(Aspose.Drawing.Color.Blue), 10, 10, 180, 80);
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // ---------------------------------------------------------------
        // 2. Build a sample RTF document that contains the image
        // ---------------------------------------------------------------
        Document rtfDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(rtfDoc);
        builder.Writeln("Sample RTF document with an image:");
        builder.InsertImage(sampleImagePath);
        string rtfPath = Path.Combine(artifactsDir, "sample.rtf");
        rtfDoc.Save(rtfPath, SaveFormat.Rtf);

        // ---------------------------------------------------------------
        // 3. Load the RTF document and extract all images
        // ---------------------------------------------------------------
        Document loadedRtf = new Document(rtfPath);
        NodeCollection shapeNodes = loadedRtf.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        // Prepare a zip archive to store the resulting TIFF files
        string zipPath = Path.Combine(artifactsDir, "ImagesArchive.zip");
        using (FileStream zipStream = new FileStream(zipPath, FileMode.Create))
        using (ZipArchive zip = new ZipArchive(zipStream, ZipArchiveMode.Create))
        {
            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage)
                    continue;

                // -------------------------------------------------------
                // 4. Convert the extracted image to a losslessly compressed TIFF
                // -------------------------------------------------------
                // Retrieve the raw image bytes from the shape
                byte[] imageBytes = shape.ImageData.ToByteArray();

                // Create a temporary document that contains only this image
                Document tempDoc = new Document();
                DocumentBuilder tempBuilder = new DocumentBuilder(tempDoc);
                tempBuilder.InsertImage(imageBytes);

                // Configure TIFF save options with lossless LZW compression
                ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
                {
                    TiffCompression = TiffCompression.Lzw
                };

                // Save the temporary document as a TIFF file (in memory)
                string tiffFileName = $"ExtractedImage_{imageIndex}.tiff";
                string tiffFullPath = Path.Combine(artifactsDir, tiffFileName);
                tempDoc.Save(tiffFullPath, tiffOptions);

                // Add the TIFF file to the zip archive
                zip.CreateEntryFromFile(tiffFullPath, tiffFileName);

                // Clean up the temporary TIFF file
                File.Delete(tiffFullPath);

                imageIndex++;
            }

            // Validation: ensure at least one image was added to the archive
            if (imageIndex == 0)
                throw new InvalidOperationException("No images were extracted from the RTF document.");
        }

        Console.WriteLine($"Extraction complete. {imageIndex} image(s) archived to: {zipPath}");
    }
}
