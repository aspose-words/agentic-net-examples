using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class ReplaceOleObjectExample
{
    public static void Main()
    {
        // Create a new blank document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a placeholder OLE object (a simple text file) into the document.
        builder.Writeln("Original OLE object:");
        using (MemoryStream textStream = new MemoryStream())
        using (StreamWriter writer = new StreamWriter(textStream))
        {
            writer.Write("This is the original OLE object content.");
            writer.Flush();
            textStream.Position = 0; // Reset stream position before insertion.

            // Insert the OLE object using the stream. ProgId "Package" creates a generic OLE package.
            builder.InsertOleObject(textStream, "Package", false, null);
        }

        // Save the document that contains the original OLE object.
        const string originalPath = "Original.docx";
        doc.Save(originalPath);

        // Load the document we just saved.
        Document loadedDoc = new Document(originalPath);
        DocumentBuilder loadedBuilder = new DocumentBuilder(loadedDoc);

        // Find the first OLE object shape in the document.
        Shape oleShape = loadedDoc.GetChildNodes(NodeType.Shape, true)
                                 .OfType<Shape>()
                                 .FirstOrDefault(s => s.ShapeType == ShapeType.OleObject);

        // Remove the existing OLE object if it was found.
        oleShape?.Remove();

        // Prepare a simple PNG image (1x1 pixel) encoded in Base64.
        const string pngBase64 =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(pngBase64);

        // Insert the new image as an OLE object at the end of the document.
        loadedBuilder.Writeln("\nReplaced OLE object (image):");
        using (MemoryStream imageStream = new MemoryStream(pngBytes))
        {
            // Insert the image stream as a generic OLE package.
            loadedBuilder.InsertOleObject(imageStream, "Package", false, null);
        }

        // Save the document after replacement.
        const string replacedPath = "Replaced.docx";
        loadedDoc.Save(replacedPath);
    }
}
