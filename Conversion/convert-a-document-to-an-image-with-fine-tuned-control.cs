using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DocumentToImageExample
{
    public class DocumentToImageConverter
    {
        /// <summary>
        /// Converts a Word document to an image with fine‑tuned rendering options.
        /// </summary>
        /// <param name="inputPath">Full path to the source .docx/.doc/.pdf file.</param>
        /// <param name="outputPath">Full path where the rendered image will be saved.</param>
        public static void ConvertToImage(string inputPath, string outputPath)
        {
            // Load the document.
            Document doc = new Document(inputPath);

            // Configure image‑saving options.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
            {
                Resolution = 300,                     // DPI – higher = better quality.
                ImageSize = new Size(1200, 1600),     // Exact pixel dimensions (optional).
                PageSet = new PageSet(0),             // Render only the first page (zero‑based).
                UseHighQualityRendering = true,      // Better visual fidelity.
                UseAntiAliasing = true,               // Smooth edges.
                PaperColor = Color.Transparent        // Transparent background.
            };

            // Save the rendered page as an image.
            doc.Save(outputPath, options);
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Expect two arguments: input file path and output image path.
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: DocumentToImageExample <inputPath> <outputPath>");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];

            try
            {
                DocumentToImageConverter.ConvertToImage(inputPath, outputPath);
                Console.WriteLine($"Document converted successfully: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }
    }
}
