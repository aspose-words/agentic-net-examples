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
        const string sampleImagePath = "sample.tif";
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                g.FillRectangle(Brushes.Black, 50, 50, 100, 100);
            }
            bitmap.Save(sampleImagePath, ImageFormat.Tiff);
        }

        // Step 2: Create an RTF document and embed the TIFF image.
        const string rtfPath = "input.rtf";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(rtfPath);

        // Step 3: Load the RTF document and extract TIFF images.
        Document loadedDoc = new Document(rtfPath);
        var shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                                 .Cast<Shape>()
                                 .Where(s => s.HasImage)
                                 .ToList();

        if (shapeNodes.Count == 0)
            throw new InvalidOperationException("No images were found in the document.");

        // Prepare the ZIP archive.
        const string zipPath = "compressed_images.zip";
        using (FileStream zipStream = new FileStream(zipPath, FileMode.Create))
        using (ZipArchive archive = new ZipArchive(zipStream, ZipArchiveMode.Update))
        {
            int index = 0;
            foreach (Shape shape in shapeNodes)
            {
                // Save the original image to a memory stream.
                using (MemoryStream originalStream = new MemoryStream())
                {
                    shape.ImageData.Save(originalStream);
                    originalStream.Position = 0;

                    // Load the image into a Bitmap.
                    using (Bitmap imgBitmap = new Bitmap(originalStream))
                    {
                        // Prepare encoder parameters for lossless LZW compression.
                        ImageCodecInfo tiffCodec = ImageCodecInfo.GetImageEncoders()
                            .FirstOrDefault(c => c.FormatID == ImageFormat.Tiff.Guid);
                        if (tiffCodec == null)
                            throw new InvalidOperationException("TIFF codec not found.");

                        EncoderParameters encoderParams = new EncoderParameters(1);
                        EncoderParameter compressionParam = new EncoderParameter(
                            Encoder.Compression,
                            (long)EncoderValue.CompressionLZW);
                        encoderParams.Param[0] = compressionParam;

                        // Save the compressed image to another memory stream.
                        using (MemoryStream compressedStream = new MemoryStream())
                        {
                            imgBitmap.Save(compressedStream, tiffCodec, encoderParams);
                            compressedStream.Position = 0;

                            // Add the compressed image to the ZIP archive.
                            string entryName = $"image-{index}.tif";
                            ZipArchiveEntry entry = archive.CreateEntry(entryName);
                            using (Stream entryStream = entry.Open())
                            {
                                compressedStream.CopyTo(entryStream);
                            }
                        }
                    }
                }
                index++;
            }
        }

        // Validation: ensure the ZIP file was created and contains entries.
        if (!File.Exists(zipPath))
            throw new InvalidOperationException("The ZIP archive was not created.");

        using (FileStream zipCheck = new FileStream(zipPath, FileMode.Open))
        using (ZipArchive archiveCheck = new ZipArchive(zipCheck, ZipArchiveMode.Read))
        {
            if (archiveCheck.Entries.Count == 0)
                throw new InvalidOperationException("No images were added to the ZIP archive.");
        }

        // Cleanup temporary files.
        File.Delete(sampleImagePath);
        File.Delete(rtfPath);
    }
}
