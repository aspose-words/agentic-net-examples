using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace AsposeWordsImagePlaceholder
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source DOCX file.
            string inputPath = @"C:\Docs\SourceDocument.docx";

            // Path where the modified document will be saved.
            string outputPath = @"C:\Docs\SourceDocument_WithPlaceholders.docx";

            // Load the existing document (lifecycle rule: load).
            Document doc = new Document(inputPath);

            // Get all Shape nodes in the document (including those inside headers/footers).
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

            // Iterate over a copy of the collection because we will modify the document structure.
            foreach (Shape shape in shapes.OfType<Shape>())
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

            // Save the modified document (lifecycle rule: save).
            doc.Save(outputPath);
        }
    }
}
