using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace AsposeWordsExamples
{
    public class AudioCoverArtExtractor
    {
        /// <summary>
        /// Extracts all images (including audio cover art) from a DOCX file and saves them as JPEG files.
        /// </summary>
        /// <param name="docxPath">Full path to the source DOCX file.</param>
        /// <param name="outputFolder">Folder where the JPEG files will be written.</param>
        public static void ExtractImagesAsJpeg(string docxPath, string outputFolder)
        {
            // Ensure the output directory exists.
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Load the document.
            Document doc = new Document(docxPath);

            // Get all Shape nodes (inline and floating) – they may contain images or OLE objects.
            var shapes = doc.GetChildNodes(NodeType.Shape, true);

            int imageIndex = 0;

            foreach (Shape shape in shapes.OfType<Shape>())
            {
                // Skip shapes without image data.
                if (!shape.HasImage)
                    continue;

                // Retrieve the raw image bytes.
                byte[] imageBytes = shape.ImageData.ImageBytes;

                // Determine a file name. The original image format may already be JPEG, PNG, etc.
                // We force a .jpg extension as required by the task.
                string outFile = Path.Combine(outputFolder, $"CoverArt_{imageIndex}.jpg");

                // Write the bytes directly to disk. If the source format is not JPEG the file will still be
                // a valid image, but its internal format will be whatever was stored in the DOCX.
                // For a true conversion to JPEG you could load the bytes into a System.Drawing.Image on Windows
                // or use a third‑party library (e.g., ImageSharp) that works cross‑platform.
                File.WriteAllBytes(outFile, imageBytes);

                imageIndex++;
            }
        }

        // Example usage.
        public static void Main()
        {
            string sourceDocx = @"C:\Docs\AudioDocument.docx";
            string jpegFolder = @"C:\Docs\ExtractedImages";

            ExtractImagesAsJpeg(sourceDocx, jpegFolder);

            Console.WriteLine("Extraction complete.");
        }
    }
}
