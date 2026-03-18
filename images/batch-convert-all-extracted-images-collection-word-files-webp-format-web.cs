using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

namespace WordImageBatchConverter
{
    class Program
    {
        /// <summary>
        /// Entry point for the console application.
        /// Expected arguments: <c>outputFolder docPath1 [docPath2] ...</c>
        /// </summary>
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: WordImageBatchConverter <outputFolder> <docPath1> [docPath2] ...");
                return;
            }

            string outputFolder = args[0];
            IEnumerable<string> docPaths = args.Skip(1);

            WebpBatchConverter.ConvertImagesToWebp(docPaths, outputFolder);
            Console.WriteLine("Image conversion completed.");
        }
    }

    public static class WebpBatchConverter
    {
        /// <summary>
        /// Converts every image found in the supplied Word documents to WebP format.
        /// Each extracted image is saved as a separate WebP file in the <paramref name="outputFolder"/>.
        /// </summary>
        /// <param name="wordFilePaths">Full paths of the Word documents to process.</param>
        /// <param name="outputFolder">Folder where the resulting WebP files will be written.</param>
        public static void ConvertImagesToWebp(IEnumerable<string> wordFilePaths, string outputFolder)
        {
            // Ensure the output directory exists.
            Directory.CreateDirectory(outputFolder);

            int globalImageIndex = 0;

            foreach (string docPath in wordFilePaths)
            {
                // Load the source document.
                Document sourceDoc = new Document(docPath);

                // Collect all shape nodes (including pictures) from the document.
                NodeCollection shapes = sourceDoc.GetChildNodes(NodeType.Shape, true);

                foreach (Shape shape in shapes.OfType<Shape>())
                {
                    if (!shape.HasImage)
                        continue;

                    // Create a temporary one‑page document that will hold the extracted image.
                    Document tempDoc = new Document();
                    DocumentBuilder builder = new DocumentBuilder(tempDoc);

                    // Insert the image bytes into the temporary document.
                    using (MemoryStream imgStream = new MemoryStream(shape.ImageData.ImageBytes))
                    {
                        builder.InsertImage(imgStream);
                    }

                    // Prepare save options to render the page as WebP.
                    ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.WebP)
                    {
                        // Only the first (and only) page is needed.
                        PageSet = new PageSet(0)
                    };

                    // Build a unique file name for the WebP image.
                    string webpFileName = Path.Combine(
                        outputFolder,
                        $"Image_{globalImageIndex}{FileFormatUtil.ImageTypeToExtension(ImageType.WebP)}");

                    // Save the temporary document as a WebP image.
                    tempDoc.Save(webpFileName, saveOptions);

                    globalImageIndex++;
                }
            }
        }
    }
}
