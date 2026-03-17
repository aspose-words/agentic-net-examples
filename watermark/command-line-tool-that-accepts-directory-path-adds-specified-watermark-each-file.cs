using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace WatermarkBatchTool
{
    class Program
    {
        /// <summary>
        /// Entry point of the command‑line tool.
        /// Usage:
        ///   WatermarkBatchTool.exe <directoryPath> [textWatermark] [imageWatermarkPath]
        /// Provide either a text watermark or an image watermark (or both; text will be applied first).
        /// </summary>
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Insufficient arguments.");
                Console.WriteLine("Usage: WatermarkBatchTool.exe <directoryPath> [textWatermark] [imageWatermarkPath]");
                return;
            }

            string directoryPath = args[0];
            string textWatermark = args.Length > 1 ? args[1] : null;
            string imageWatermarkPath = args.Length > 2 ? args[2] : null;

            if (!Directory.Exists(directoryPath))
            {
                Console.WriteLine($"Directory does not exist: {directoryPath}");
                return;
            }

            // Process supported Word document extensions.
            string[] supportedExtensions = new[] { ".doc", ".docx", ".rtf", ".odt", ".xml", ".mht", ".pdf" };
            string[] files = Directory.GetFiles(directoryPath, "*.*", SearchOption.TopDirectoryOnly);
            foreach (string filePath in files)
            {
                if (Array.IndexOf(supportedExtensions, Path.GetExtension(filePath).ToLowerInvariant()) < 0)
                    continue; // Skip unsupported files.

                try
                {
                    // Load the document.
                    Document doc = new Document(filePath);

                    // Apply text watermark if supplied.
                    if (!string.IsNullOrWhiteSpace(textWatermark))
                    {
                        // Use default options; can be customized via TextWatermarkOptions if needed.
                        doc.Watermark.SetText(textWatermark);
                    }

                    // Apply image watermark if supplied.
                    if (!string.IsNullOrWhiteSpace(imageWatermarkPath) && File.Exists(imageWatermarkPath))
                    {
                        // Configure image watermark options (example: scale = 5, no washout).
                        ImageWatermarkOptions imgOptions = new ImageWatermarkOptions
                        {
                            Scale = 5,
                            IsWashout = false
                        };
                        doc.Watermark.SetImage(imageWatermarkPath, imgOptions);
                    }

                    // Overwrite the original file with the watermarked version.
                    doc.Save(filePath);
                    Console.WriteLine($"Watermarked: {Path.GetFileName(filePath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to process '{Path.GetFileName(filePath)}': {ex.Message}");
                }
            }

            Console.WriteLine("Processing completed.");
        }
    }
}
