using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace WatermarkBatchProcessor
{
    class Program
    {
        static void Main(string[] args)
        {
            // Example usage:
            // args[0] - folder containing .doc files
            // args[1] - folder where watermarked files will be saved
            // args[2] - full path to the image used as watermark
            if (args.Length != 3)
            {
                Console.WriteLine("Usage: WatermarkBatchProcessor <inputFolder> <outputFolder> <watermarkImagePath>");
                return;
            }

            string inputFolder = args[0];
            string outputFolder = args[1];
            string watermarkImagePath = args[2];

            // Ensure output folder exists
            Directory.CreateDirectory(outputFolder);

            // Process each .doc file in the input folder
            foreach (string docPath in Directory.GetFiles(inputFolder, "*.doc"))
            {
                // Load the document using the Document constructor (load rule)
                Document doc = new Document(docPath);

                // Configure watermark options (optional)
                ImageWatermarkOptions options = new ImageWatermarkOptions
                {
                    // Example: make the watermark more visible
                    IsWashout = false,
                    Scale = 5 // Adjust scale as needed
                };

                // Apply the image watermark using the Watermark.SetImage method (watermark rule)
                doc.Watermark.SetImage(watermarkImagePath, options);

                // Determine output file path (preserve original file name)
                string outputPath = Path.Combine(outputFolder, Path.GetFileName(docPath));

                // Save the modified document (save rule)
                doc.Save(outputPath);
            }

            Console.WriteLine("Watermarking completed.");
        }
    }
}
