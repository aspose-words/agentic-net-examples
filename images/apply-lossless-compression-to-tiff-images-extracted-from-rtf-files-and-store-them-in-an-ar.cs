using System;
using System.Collections.Generic;
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
        // Prepare output folders.
        string artifactsDir = "Artifacts";
        Directory.CreateDirectory(artifactsDir);
        string tempDir = Path.Combine(artifactsDir, "Temp");
        Directory.CreateDirectory(tempDir);

        // 1. Create a deterministic sample image (PNG) using Aspose.Drawing.
        string sampleImagePath = Path.Combine(tempDir, "sample.png");
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.LightBlue);
            }
            bitmap.Save(sampleImagePath);
        }

        // 2. Create a Word document, insert the image, and save it as RTF.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        string rtfPath = Path.Combine(artifactsDir, "sample.rtf");
        doc.Save(rtfPath, SaveFormat.Rtf);

        // 3. Load the RTF document.
        Document loadedDoc = new Document(rtfPath);

        // 4. Extract each image, convert it to a TIFF with lossless LZW compression, and collect the TIFF file paths.
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        List<string> tiffFiles = new List<string>();
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            // Save the original image bytes to a memory stream.
            using (MemoryStream imgStream = new MemoryStream())
            {
                shape.ImageData.Save(imgStream);
                imgStream.Position = 0;

                // Create a temporary document that contains only this image.
                Document imgDoc = new Document();
                DocumentBuilder imgBuilder = new DocumentBuilder(imgDoc);
                imgBuilder.InsertImage(imgStream);

                // Configure TIFF save options with lossless LZW compression.
                ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
                {
                    TiffCompression = TiffCompression.Lzw
                };

                // Save the image as a TIFF file.
                string tiffPath = Path.Combine(tempDir, $"image_{imageIndex}.tiff");
                imgDoc.Save(tiffPath, tiffOptions);
                tiffFiles.Add(tiffPath);
                imageIndex++;
            }
        }

        // Validate that at least one TIFF image was produced.
        if (tiffFiles.Count == 0)
            throw new InvalidOperationException("No images were extracted from the RTF document.");

        // 5. Store all TIFF images in a ZIP archive.
        string zipPath = Path.Combine(artifactsDir, "ImagesArchive.zip");
        using (FileStream zipToOpen = new FileStream(zipPath, FileMode.Create))
        {
            using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Create))
            {
                foreach (string tiffFile in tiffFiles)
                {
                    archive.CreateEntryFromFile(tiffFile, Path.GetFileName(tiffFile));
                }
            }
        }

        // Validate that the ZIP archive was created.
        if (!File.Exists(zipPath))
            throw new InvalidOperationException("Failed to create the ZIP archive.");
    }
}
