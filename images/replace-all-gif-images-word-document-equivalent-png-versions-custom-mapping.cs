using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
            // Load the document.
            Document doc = new Document(inputPath);

            int gifCounter = 0;

            // Iterate over all Shape nodes (including inline and floating images).
            foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>())
            {
                if (shape.HasImage && shape.ImageData.ImageType == ImageType.Gif)
                {
                    if (gifIndexToPngPath.TryGetValue(gifCounter, out string pngPath) && File.Exists(pngPath))
                    {
                        shape.ImageData.SetImage(pngPath);
                    }

                    gifCounter++;
                }
            }

            // Save the modified document.
            doc.Save(outputPath);
        }

        static void Main()
        {
            // Create a temporary working folder.
            string workDir = Path.Combine(Path.GetTempPath(), "AsposeGifToPngDemo");
            Directory.CreateDirectory(workDir);

            // Paths for the sample files.
            string sourceDoc = Path.Combine(workDir, "SampleWithGifs.docx");
            string resultDoc = Path.Combine(workDir, "SampleWithGifs_Converted.docx");
            string gifPath = Path.Combine(workDir, "sample.gif");
            string pngPath = Path.Combine(workDir, "replacement.png");

            // Write a tiny 1x1 transparent GIF.
            byte[] gifBytes = Convert.FromBase64String(
                "R0lGODdhAQABAPAAAP///wAAACH5BAAAAAAALAAAAAABAAEAAAICRAEAOw==");
            File.WriteAllBytes(gifPath, gifBytes);

            // Write a tiny 1x1 transparent PNG.
            byte[] pngBytes = Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK9cAAAAASUVORK5CYII=");
            File.WriteAllBytes(pngPath, pngBytes);

            // Create a simple document and insert the GIF image.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Document with a GIF image:");
            builder.InsertImage(gifPath);
            doc.Save(sourceDoc);

            // Build the mapping: first GIF (index 0) -> our PNG replacement.
            var gifToPngMap = new Dictionary<int, string>
            {
                { 0, pngPath }
            };

            // Perform the replacement.
            ReplaceGifWithPng(sourceDoc, resultDoc, gifToPngMap);

            Console.WriteLine("GIF images have been replaced and the document saved to:");
            Console.WriteLine(resultDoc);
        }
    }
}
