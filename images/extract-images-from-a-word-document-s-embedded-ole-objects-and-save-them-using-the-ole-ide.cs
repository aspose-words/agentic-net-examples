using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Loading;
using Aspose.Drawing; // Provides Bitmap, Graphics, Color

public class Program
{
    public static void Main()
    {
        // Folder for generated files
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample Word document with an embedded OLE object.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create dummy data for the OLE object (e.g., a simple text file content)
        byte[] oleData = System.Text.Encoding.UTF8.GetBytes("Sample OLE content");
        using (MemoryStream oleStream = new MemoryStream(oleData))
        {
            // Insert the OLE object as an icon so that it has an image representation.
            // ProgId "Package" denotes a generic OLE package.
            builder.Writeln("Below is an embedded OLE object displayed as an icon:");
            Shape oleShape = builder.InsertOleObject(oleStream, "Package", true, null);
            // Ensure the shape is added to the document.
            builder.InsertParagraph();
        }

        // Save the document to disk.
        string docPath = Path.Combine(outputDir, "SampleDocument.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 2. Load the document and extract images from embedded OLE objects.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // Identify OLE objects.
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat == null)
                continue; // Not an OLE object.

            // If the OLE shape has an image (icon), extract it.
            if (shape.HasImage)
            {
                // Determine file extension based on the image type.
                string extension = Aspose.Words.FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                // Build a deterministic file name using the OLE ProgId.
                string imageFileName = $"OleImage_{oleFormat.ProgId}_{imageIndex}{extension}";
                string imagePath = Path.Combine(outputDir, imageFileName);

                // Save the image.
                shape.ImageData.Save(imagePath);
                imageIndex++;
            }
        }

        // Validate that at least one image was extracted.
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from OLE objects.");

        // -----------------------------------------------------------------
        // 3. (Optional) Extract the raw OLE data itself for reference.
        // -----------------------------------------------------------------
        int oleDataIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat == null)
                continue;

            // Use the suggested file extension for the embedded object.
            string oleExtension = oleFormat.SuggestedExtension ?? ".bin";
            string oleFileName = $"OleData_{oleFormat.ProgId}_{oleDataIndex}{oleExtension}";
            string olePath = Path.Combine(outputDir, oleFileName);

            // Save the OLE data to a file.
            oleFormat.Save(olePath);
            oleDataIndex++;
        }

        // The example finishes without requiring user interaction.
    }
}
