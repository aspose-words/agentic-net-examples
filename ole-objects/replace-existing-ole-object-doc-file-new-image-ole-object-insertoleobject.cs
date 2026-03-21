using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class ReplaceOleObject
{
    static void Main()
    {
        // Create temporary files and folders.
        string tempDir = Path.GetTempPath();
        string sourceDocPath = Path.Combine(tempDir, "Original.docx");
        string outputDocPath = Path.Combine(tempDir, "Modified.docx");
        string textFilePath = Path.Combine(tempDir, "sample.txt");
        string newImagePath = Path.Combine(tempDir, "Replacement.png");

        // Prepare a simple text file to embed as an OLE object.
        File.WriteAllText(textFilePath, "Sample text for OLE object.");

        // Prepare a minimal PNG image (1x1 pixel, transparent).
        byte[] pngBytes = new byte[]
        {
            0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
            0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
            0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
            0x08,0x06,0x00,0x00,0x00,0x1F,0x15,0xC4,
            0x89,0x00,0x00,0x00,0x0A,0x49,0x44,0x41,
            0x54,0x78,0x9C,0x63,0x60,0x00,0x00,0x00,
            0x02,0x00,0x01,0xE2,0x21,0xBC,0x33,0x00,
            0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,
            0x42,0x60,0x82
        };
        File.WriteAllBytes(newImagePath, pngBytes);

        // Create a new document and embed the text file as an OLE object.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        using (FileStream txtStream = File.OpenRead(textFilePath))
        {
            builder.InsertOleObject(txtStream, "Package", false, null);
        }
        doc.Save(sourceDocPath);

        // Reload the document to simulate working with an existing file.
        doc = new Document(sourceDocPath);

        // Locate the first shape that contains an OLE object.
        Shape oleShape = null;
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes)
        {
            if (shape.ShapeType == ShapeType.OleObject)
            {
                oleShape = shape;
                break;
            }
        }

        if (oleShape == null)
        {
            Console.WriteLine("No OLE object found in the document.");
            return;
        }

        // Move the builder to the position of the OLE shape before removing it.
        DocumentBuilder replaceBuilder = new DocumentBuilder(doc);
        replaceBuilder.MoveTo(oleShape);
        oleShape.Remove();

        // Insert the new image as an OLE object.
        using (FileStream imageStream = File.OpenRead(newImagePath))
        {
            replaceBuilder.InsertOleObject(imageStream, "Package", false, null);
        }

        // Save the modified document.
        doc.Save(outputDocPath);

        Console.WriteLine($"Source document: {sourceDocPath}");
        Console.WriteLine($"Modified document: {outputDocPath}");
    }
}
