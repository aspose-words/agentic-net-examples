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
        string inputPath = Path.Combine(dataDir, "Barcodes.docx");

        // Load the document (creation and loading follow the provided lifecycle rules).
        Document doc = new Document(inputPath);

        // Update all fields so that barcode fields are rendered as images.
        doc.UpdateFields();

        // Iterate through all shapes in the document and extract those that are images.
        int imageIndex = 0;
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            if (shape.IsImage)
            {
                // Build a unique file name for each extracted barcode image.
                string outputPath = Path.Combine(dataDir, $"BarcodeImage_{imageIndex}.png");

                // Save the image data to the file system.
                shape.ImageData.Save(outputPath);

                imageIndex++;
            }
        }

        // Optionally save the updated document (still using the provided save rule).
        string updatedDocPath = Path.Combine(dataDir, "Barcodes_Updated.docx");
        doc.Save(updatedDocPath);
    }
}
