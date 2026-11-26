using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample TIFF image.
        const string sampleTiffPath = "sample.tif";
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.LightBlue);
                g.DrawRectangle(new Pen(Color.DarkBlue, 5), 20, 20, 160, 160);
            }
            bitmap.Save(sampleTiffPath, ImageFormat.Tiff);
        }

        // Step 2: Create an RTF document and insert the TIFF image.
        const string rtfPath = "sample.rtf";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleTiffPath);
        doc.Save(rtfPath, SaveFormat.Rtf);

        // Step 3: Load the RTF document and extract embedded images.
        Document loadedDoc = new Document(rtfPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        var extractedPaths = new System.Collections.Generic.List<string>();

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage) continue;

            // Save the extracted image to a temporary file.
            string extractedPath = $"extracted_{imageIndex}.tif";
            shape.ImageData.Save(extractedPath);
            extractedPaths.Add(extractedPath);
            imageIndex++;
        }

        if (extractedPaths.Count == 0)
            throw new InvalidOperationException("No images were extracted from the RTF document.");

        // Step 4: Apply lossless compression (LZW) to each extracted TIFF image.
        var compressedPaths = new System.Collections.Generic.List<string>();
        ImageCodecInfo tiffCodec = ImageCodecInfo.GetImageEncoders()
            .FirstOrDefault(c => c.FormatID == ImageFormat.Tiff.Guid);
        if (tiffCodec == null)
            throw new InvalidOperationException("TIFF codec not found.");

        Encoder compressionEncoder = Encoder.Compression;
        EncoderParameters encoderParams = new EncoderParameters(1);
        encoderParams.Param[0] = new EncoderParameter(compressionEncoder, (long)EncoderValue.CompressionLZW);

        for (int i = 0; i < extractedPaths.Count; i++)
        {
            string srcPath = extractedPaths[i];
            string compressedPath = $"compressed_{i}.tif";

            using (FileStream fs = new FileStream(srcPath, FileMode.Open, FileAccess.Read))
            using (MemoryStream ms = new MemoryStream())
            {
                fs.CopyTo(ms);
                ms.Position = 0;

                using (Bitmap bmp = new Bitmap(ms))
                {
                    bmp.Save(compressedPath, tiffCodec, encoderParams);
                }
            }

            compressedPaths.Add(compressedPath);
        }

        // Step 5: Store compressed images in a ZIP archive.
        const string archivePath = "ImagesArchive.zip";
        using (FileStream zipToOpen = new FileStream(archivePath, FileMode.Create))
        using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Update))
        {
            foreach (string path in compressedPaths)
            {
                ZipArchiveEntry entry = archive.CreateEntry(Path.GetFileName(path));
                using (Stream entryStream = entry.Open())
                using (FileStream fileStream = new FileStream(path, FileMode.Open, FileAccess.Read))
                {
                    fileStream.CopyTo(entryStream);
                }
            }
        }

        // Validation: ensure the archive contains entries.
        using (ZipArchive archive = ZipFile.OpenRead(archivePath))
        {
            if (!archive.Entries.Any())
                throw new InvalidOperationException("The archive does not contain any images.");
        }

        // Cleanup temporary files (optional).
        foreach (var path in extractedPaths) File.Delete(path);
        foreach (var path in compressedPaths) File.Delete(path);
        File.Delete(sampleTiffPath);
        File.Delete(rtfPath);
    }
}
