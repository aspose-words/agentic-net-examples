using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace TiffConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            // Expect three arguments: input file path, DPI, and TIFF compression type.
            if (args.Length != 3)
            {
                Console.WriteLine("Usage: TiffConverter <inputPath> <dpi> <compression>");
                Console.WriteLine("Compression values: None, Rle, Lzw, Ccitt3, Ccitt4");
                return;
            }

            string inputPath = args[0];
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' does not exist.");
                return;
            }

            // Parse DPI (dots per inch). Must be a positive integer.
            if (!int.TryParse(args[1], out int dpi) || dpi <= 0)
            {
                Console.WriteLine("Error: DPI must be a positive integer.");
                return;
            }

            // Parse compression type into the TiffCompression enum.
            if (!Enum.TryParse<TiffCompression>(args[2], true, out TiffCompression compression))
            {
                Console.WriteLine("Error: Invalid compression type. Use one of: None, Rle, Lzw, Ccitt3, Ccitt4.");
                return;
            }

            // Load the Word document.
            Document doc = new Document(inputPath);

            // Configure image save options for TIFF.
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
            {
                // Apply the requested DPI to both dimensions.
                HorizontalResolution = dpi,
                VerticalResolution = dpi,

                // Set the requested compression scheme.
                TiffCompression = compression,

                // Render each page as a separate frame in a multi‑page TIFF.
                PageLayout = MultiPageLayout.TiffFrames()
            };

            // Build output file path: same directory, same name with .tiff extension.
            string outputPath = Path.ChangeExtension(inputPath, ".tiff");

            // Save the document as a TIFF image using the configured options.
            doc.Save(outputPath, saveOptions);

            Console.WriteLine($"Document converted successfully: {outputPath}");
        }
    }
}
