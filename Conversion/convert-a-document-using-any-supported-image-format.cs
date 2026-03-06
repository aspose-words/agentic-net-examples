using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DocumentConversionExample
{
    public class Converter
    {
        /// <summary>
        /// Converts the first page of a Word document to an image of the specified format.
        /// </summary>
        /// <param name="inputFilePath">Path to the source document (e.g., .docx, .pdf).</param>
        /// <param name="outputFilePath">Path where the resulting image will be saved.</param>
        /// <param name="imageFormat">Desired image format (e.g., SaveFormat.Png, SaveFormat.Jpeg).</param>
        public static void ConvertDocumentToImage(string inputFilePath, string outputFilePath, SaveFormat imageFormat)
        {
            // Load the document from the file system.
            Document doc = new Document(inputFilePath);

            // Create ImageSaveOptions specifying the target image format.
            ImageSaveOptions saveOptions = new ImageSaveOptions(imageFormat)
            {
                // Optional: set resolution, quality, etc.
                Resolution = 300,          // 300 DPI for high‑quality rendering.
                JpegQuality = 90          // Used only for JPEG format.
            };

            // Save the document as an image. For raster image formats only the first page is saved.
            doc.Save(outputFilePath, saveOptions);
        }

        /// <summary>
        /// Demonstrates converting a document to PNG and JPEG images using streams.
        /// </summary>
        public static void Demo()
        {
            string sourcePath = @"C:\Docs\SampleDocument.docx";

            // Convert to PNG and write to a file.
            ConvertDocumentToImage(sourcePath, @"C:\Output\SampleDocument.png", SaveFormat.Png);

            // Convert to JPEG and write to a file.
            ConvertDocumentToImage(sourcePath, @"C:\Output\SampleDocument.jpg", SaveFormat.Jpeg);

            // Convert to PNG using a memory stream (no intermediate file).
            using (MemoryStream inputStream = new MemoryStream(File.ReadAllBytes(sourcePath)))
            using (MemoryStream outputStream = new MemoryStream())
            {
                Document doc = new Document(inputStream);
                ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
                {
                    Resolution = 150
                };
                doc.Save(outputStream, options);

                // Reset stream position before using the image data.
                outputStream.Position = 0;
                File.WriteAllBytes(@"C:\Output\SampleDocument_Stream.png", outputStream.ToArray());
            }
        }
    }

    public class Program
    {
        /// <summary>
        /// Entry point required for console execution.
        /// If three arguments are supplied, they are interpreted as:
        ///   1. Input document path
        ///   2. Output image path
        ///   3. Image format (e.g., Png, Jpeg, Bmp, Gif, Tiff)
        /// Otherwise the built‑in demo is executed.
        /// </summary>
        public static void Main(string[] args)
        {
            if (args.Length >= 3)
            {
                string inputPath = args[0];
                string outputPath = args[1];
                if (!Enum.TryParse<SaveFormat>(args[2], true, out var format))
                {
                    Console.WriteLine("Invalid image format. Use one of: Png, Jpeg, Bmp, Gif, Tiff, etc.");
                    return;
                }

                Converter.ConvertDocumentToImage(inputPath, outputPath, format);
                Console.WriteLine($"Document converted to {format} and saved to '{outputPath}'.");
            }
            else
            {
                // No arguments – run the demonstration code.
                Converter.Demo();
            }
        }
    }
}
