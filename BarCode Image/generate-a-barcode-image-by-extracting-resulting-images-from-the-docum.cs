using Aspose.Words;
using Aspose.Words.Drawing;
using System;
using System.IO;

class Program
{
    static void Main()
    {
        // Path to the DOCX that contains barcode fields.
        string inputPath = @"C:\Docs\Barcodes.docx";

        // Folder where extracted barcode images will be saved.
        string outputFolder = @"C:\Docs\ExtractedImages";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Load the document.
        Document doc = new Document(inputPath);

        // Update all fields so that barcode fields are rendered as images.
        doc.UpdateFields();

        // Get all Shape nodes (they hold images) from the document.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapeNodes)
        {
            // Only process shapes that actually contain an image.
            if (shape.HasImage)
            {
                // Create a unique file name for each extracted image.
                string imagePath = Path.Combine(outputFolder, $"BarcodeImage_{imageIndex}.png");

                // Save the image data to the file system.
                shape.ImageData.Save(imagePath);

                imageIndex++;
            }
        }

        // Save the updated document (optional, shows that fields were updated).
        doc.Save(Path.Combine(outputFolder, "Barcodes_Updated.docx"));
    }
}
