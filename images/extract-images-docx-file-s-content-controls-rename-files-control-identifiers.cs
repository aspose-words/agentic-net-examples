using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Markup;

namespace Example
{
    public static class ContentControlImageExtractor
    {
        /// <summary>
        /// Extracts all images that are located inside Structured Document Tags (content controls)
        /// from the specified DOCX file and saves them to the output folder.
        /// Each image file name is built from the content control identifier (Id) and a running index.
        /// </summary>
        /// <param name="inputDocxPath">Full path to the source DOCX file.</param>
        /// <param name="outputFolderPath">Folder where extracted images will be saved.</param>
        public static void ExtractImagesFromContentControls(string inputDocxPath, string outputFolderPath)
        {
            // Ensure the output directory exists.
            if (!Directory.Exists(outputFolderPath))
                Directory.CreateDirectory(outputFolderPath);

            // Load the document.
            Document doc = new Document(inputDocxPath);

            // Get all Structured Document Tags (content controls).
            StructuredDocumentTagCollection sdtCollection = doc.Range.StructuredDocumentTags;

            foreach (IStructuredDocumentTag sdt in sdtCollection)
            {
                // Prefer the Tag property; otherwise fall back to the numeric Id.
                string controlId = !string.IsNullOrEmpty(sdt.Tag) ? sdt.Tag : sdt.Id.ToString();

                // Find all Shape nodes that are descendants of this content control.
                NodeCollection shapeNodes = sdt.GetChildNodes(NodeType.Shape, true);

                int imageIndex = 0;
                foreach (Shape shape in shapeNodes.OfType<Shape>())
                {
                    if (!shape.HasImage)
                        continue;

                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string fileName = $"Image_{controlId}_{imageIndex}{extension}";
                    string fullPath = Path.Combine(outputFolderPath, fileName);

                    shape.ImageData.Save(fullPath);
                    imageIndex++;
                }
            }
        }
    }

    class Program
    {
        /// <summary>
        /// Entry point required by the C# compiler. Demonstrates how to call the extractor.
        /// Adjust the paths below to point to an existing DOCX file and a writable output folder.
        /// </summary>
        static void Main(string[] args)
        {
            // Example usage – replace with real paths or pass them via command‑line arguments.
            string inputDocx = args.Length > 0 ? args[0] : "sample.docx";
            string outputFolder = args.Length > 1 ? args[1] : "ExtractedImages";

            try
            {
                ContentControlImageExtractor.ExtractImagesFromContentControls(inputDocx, outputFolder);
                Console.WriteLine($"Images extracted to: {Path.GetFullPath(outputFolder)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
