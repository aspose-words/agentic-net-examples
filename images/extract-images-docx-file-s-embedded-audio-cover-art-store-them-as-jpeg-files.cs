using System;
using System.IO;
using System.Linq;
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
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            Document doc = new Document(docxPath);
            var shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapes.OfType<Shape>())
            {
                if (!shape.HasImage)
                    continue;

                byte[] imageBytes = shape.ImageData.ImageBytes;
                string outFile = Path.Combine(outputFolder, $"CoverArt_{imageIndex}.jpg");
                File.WriteAllBytes(outFile, imageBytes);
                imageIndex++;
            }
        }

        private static byte[] GetPlaceholderImageBytes()
        {
            // 1x1 transparent PNG
            const string base64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
            return Convert.FromBase64String(base64);
        }

        public static void Main()
        {
            string sourceDocx = Path.Combine(Directory.GetCurrentDirectory(), "SampleAudioDocument.docx");
            if (!File.Exists(sourceDocx))
            {
                Document doc = new Document();
                DocumentBuilder builder = new DocumentBuilder(doc);
                builder.Writeln("Sample document containing an image (simulating cover art).");
                builder.InsertImage(GetPlaceholderImageBytes());
                doc.Save(sourceDocx);
            }

            string jpegFolder = Path.Combine(Directory.GetCurrentDirectory(), "ExtractedImages");
            ExtractImagesAsJpeg(sourceDocx, jpegFolder);

            Console.WriteLine("Extraction complete. Images saved to: " + jpegFolder);
        }
    }
}
