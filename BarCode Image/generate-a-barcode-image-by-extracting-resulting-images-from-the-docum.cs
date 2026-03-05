using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load the DOCX document that contains barcode fields.
        Document doc = new Document("Barcodes.docx");

        // Update fields to ensure barcode images are generated.
        doc.UpdateFields();

        // Get all Shape nodes (they can contain images).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapeNodes)
        {
            // Check if the shape actually holds an image.
            if (shape.IsImage)
            {
                // Save the extracted image to a file (PNG format is used here).
                string imageFileName = $"BarcodeImage_{imageIndex}.png";
                shape.ImageData.Save(imageFileName);
                imageIndex++;
            }
        }

        // Save the document after processing (optional).
        doc.Save("Barcodes_Processed.docx");
    }
}
