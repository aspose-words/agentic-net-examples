using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class BarcodeImageExtractor
{
    static void Main()
    {
        // Path to the input DOCX that contains barcode fields.
        string dataDir = @"C:\Data\";
        string inputDocPath = Path.Combine(dataDir, "Barcodes.docx");

        // Load the document.
        Document doc = new Document(inputDocPath);

        // Update all fields so that barcode fields are rendered as images.
        doc.UpdateFields();

        // Get all Shape nodes (they can contain images).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapeNodes)
        {
            // Check if the shape actually holds an image.
            if (shape.IsImage)
            {
                // Build a unique file name for each extracted image.
                string outputImagePath = Path.Combine(dataDir, $"BarcodeImage_{imageIndex}.png");

                // Save the image data to the file system.
                shape.ImageData.Save(outputImagePath);

                imageIndex++;
            }
        }

        // Optionally save the updated document (fields now contain images).
        string updatedDocPath = Path.Combine(dataDir, "Barcodes_Updated.docx");
        doc.Save(updatedDocPath);
    }
}
