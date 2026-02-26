using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class CustomizeImageAndConvertToPdf
{
    static void Main()
    {
        // Load the existing DOCX document.
        // This uses the provided Document(string) constructor (load rule).
        Document doc = new Document("InputDocument.docx");

        // Iterate through all shapes in the document and find the images.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes)
        {
            if (shape.IsImage)
            {
                // Example customization: set a fixed width and preserve aspect ratio.
                // Width is set in points (1 point = 1/72 inch).
                shape.Width = 200; // 200 points ≈ 2.78 inches
                shape.Height = 0;  // Height = 0 tells Aspose.Words to keep the original aspect ratio.
                
                // Example: rotate the image 45 degrees.
                shape.Rotation = 45;
                
                // Example: add a simple border around the image.
                shape.StrokeColor = System.Drawing.Color.Black;
                shape.StrokeWeight = 0.5; // points
            }
        }

        // Save the modified document as PDF.
        // This uses the provided Document.Save(string) overload which determines format from the extension.
        doc.Save("OutputDocument.pdf");
    }
}
