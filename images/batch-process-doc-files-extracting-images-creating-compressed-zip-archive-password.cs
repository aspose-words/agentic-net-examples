using System;
using System.IO;
using System.IO.Compression;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace AsposeWordsBatchImageExtractor
{
    class Program
    {
        static void Main(string[] args)
        {
            // Expected arguments:
            // args[0] - folder containing DOC/DOCX files
            // args[1] - output ZIP file path
            // args[2] - (optional) password for ZIP encryption (not supported by built‑in .NET ZipArchive)
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: AsposeWordsBatchImageExtractor <sourceFolder> <outputZipPath> [zipPassword]");
                return;
            }

            string sourceFolder = args[0];
            string zipPath = args[1];
            // Make the password variable nullable to avoid CS8600 warning.
            string? zipPassword = args.Length > 2 ? args[2] : null; // Placeholder – .NET ZipArchive has no password support

            // Create the ZIP archive
            using (FileStream zipStream = new FileStream(zipPath, FileMode.Create))
            using (ZipArchive archive = new ZipArchive(zipStream, ZipArchiveMode.Create))
            {
                // Process each Word document in the folder (including .doc and .docx)
                foreach (string docFile in Directory.GetFiles(sourceFolder, "*.doc*"))
                {
                    // Load the document using Aspose.Words
                    Document doc = new Document(docFile);

                    // Collect all Shape nodes (including images)
                    NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
                    int imageIndex = 0;

                    foreach (Shape shape in shapes)
                    {
                        if (!shape.IsImage)
                            continue; // Skip non‑image shapes

                        // Extract raw image bytes
                        byte[] imageBytes = shape.ImageData.ImageBytes;

                        // Determine a suitable file extension for the image type
                        string extension = GetImageExtension(shape.ImageData.ImageType);

                        // Build a unique entry name: <DocumentName>_img<index>.<ext>
                        string entryName = $"{Path.GetFileNameWithoutExtension(docFile)}_img{imageIndex}{extension}";
                        imageIndex++;

                        // Add the image to the ZIP archive (compression level set to Optimal)
                        ZipArchiveEntry entry = archive.CreateEntry(entryName, CompressionLevel.Optimal);
                        using (Stream entryStream = entry.Open())
                        {
                            entryStream.Write(imageBytes, 0, imageBytes.Length);
                        }
                    }
                }

                // NOTE: The built‑in ZipArchive does not provide password protection.
                // If password‑protected ZIPs are required, integrate a third‑party library
                // such as DotNetZip or SharpZipLib and replace the archive creation logic
                // accordingly, using the 'zipPassword' variable.
                if (!string.IsNullOrEmpty(zipPassword))
                {
                    // Placeholder for password handling with a third‑party ZIP library.
                }
            }

            Console.WriteLine($"Image extraction complete. Archive saved to: {zipPath}");
        }

        /// <summary>
        /// Maps Aspose.Words.Drawing.ImageType to a conventional file extension.
        /// </summary>
        private static string GetImageExtension(ImageType imageType)
        {
            // Most common image types are covered; default to .bin for unknown types.
            switch (imageType)
            {
                case ImageType.Jpeg:
                    return ".jpg";
                case ImageType.Png:
                    return ".png";
                case ImageType.Bmp:
                    return ".bmp";
                case ImageType.Gif:
                    return ".gif";
                case ImageType.Emf:
                    return ".emf";
                case ImageType.Wmf:
                    return ".wmf";
                // Aspose.Words versions prior to 23.5 do not expose Tiff or Icon enum members.
                // They are safely omitted; unknown types will fall through to the default case.
                default:
                    return ".bin";
            }
        }
    }
}
