using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

namespace AsposeWordsImagePlaceholder
{
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output paths relative to the current directory.
            string currentDir = Directory.GetCurrentDirectory();
            string inputPath = Path.Combine(currentDir, "SourceDocument.docx");
            string outputPath = Path.Combine(currentDir, "SourceDocument_WithPlaceholders.docx");

            // Create a sample document with an image if it does not already exist.
            if (!File.Exists(inputPath))
            {
                // Create a new blank document.
                Document sampleDoc = new Document();
                DocumentBuilder builder = new DocumentBuilder(sampleDoc);

                // Insert a simple 1x1 pixel PNG image from a base64 string.
                const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2V8AAAAASUVORK5CYII=";
                byte[] pngBytes = Convert.FromBase64String(base64Png);
                using (MemoryStream imgStream = new MemoryStream(pngBytes))
                {
                    builder.InsertImage(imgStream);
                }

                // Save the sample document.
                sampleDoc.Save(inputPath);
            }

            // Load the existing document.
            Document doc = new Document(inputPath);

            // Get all Shape nodes in the document (including those inside headers/footers).
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

            // Iterate over a copy of the collection because we will modify the document structure.
            foreach (Shape shape in shapeNodes.OfType<Shape>().ToList())
            {
                // Process only shapes that actually contain an image.
                if (shape.HasImage)
                {
                    // Create a Run node that will act as the placeholder text.
                    Run placeholder = new Run(doc, "[Image]");

                    // Insert the placeholder after the image shape.
                    shape.ParentNode.InsertAfter(placeholder, shape);

                    // Remove the original image shape from the document.
                    shape.Remove();
                }
            }

            // Save the modified document.
            doc.Save(outputPath);
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
