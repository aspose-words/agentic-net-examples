using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace AsposeWordsGifToPng
{
    class Program
    {
        /// <summary>
        /// Replaces every GIF image in the source document with a PNG image according to a user‑provided mapping.
        /// </summary>
        /// <param name="inputPath">Full path to the source .docx/.doc file.</param>
        /// <param name="outputPath">Full path where the modified document will be saved.</param>
        /// <param name="gifIndexToPngPath">
        /// Mapping where the key is the zero‑based index of a GIF image encountered in the document
        /// (order of appearance) and the value is the full path to the replacement PNG file.
        /// </param>
        static void ReplaceGifWithPng(string inputPath, string outputPath, Dictionary<int, string> gifIndexToPngPath)
        {
            // Load the document using the standard constructor (load rule).
            Document doc = new Document(inputPath);

            // Counter for GIF images found so far.
            int gifCounter = 0;

            // Iterate over all Shape nodes in the document (including inline and floating images).
            foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>())
            {
                // Process only shapes that actually contain an image and whose image type is GIF.
                if (shape.HasImage && shape.ImageData.ImageType == ImageType.Gif)
                {
                    // Try to obtain a replacement PNG path for the current GIF index.
                    if (gifIndexToPngPath.TryGetValue(gifCounter, out string pngPath) && File.Exists(pngPath))
                    {
                        // Replace the image data with the PNG file (replace rule).
                        shape.ImageData.SetImage(pngPath);
                    }

                    // Move to the next GIF image.
                    gifCounter++;
                }
            }

            // Save the modified document using the standard Save method (save rule).
            doc.Save(outputPath);
        }

        static void Main()
        {
            // Example usage:

            // Path to the original Word document.
            string sourceDoc = @"C:\Docs\SampleWithGifs.docx";

            // Path where the updated document will be written.
            string resultDoc = @"C:\Docs\SampleWithGifs_Converted.docx";

            // Build a custom mapping: GIF #0 -> PNG file A, GIF #1 -> PNG file B, etc.
            var gifToPngMap = new Dictionary<int, string>
            {
                { 0, @"C:\Images\Replacement0.png" },
                { 1, @"C:\Images\Replacement1.png" },
                // Add more entries as needed.
            };

            ReplaceGifWithPng(sourceDoc, resultDoc, gifToPngMap);

            Console.WriteLine("GIF images have been replaced and the document saved to:");
            Console.WriteLine(resultDoc);
        }
    }
}
