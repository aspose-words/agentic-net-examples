using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Define temporary file paths.
        string tempFolder = Path.Combine(Path.GetTempPath(), "AsposeOleDemo");
        Directory.CreateDirectory(tempFolder);
        string textFilePath = Path.Combine(tempFolder, "sample.txt");
        string originalDocPath = Path.Combine(tempFolder, "Original.doc");
        string modifiedDocPath = Path.Combine(tempFolder, "Modified.doc");
        string imageFilePath = Path.Combine(tempFolder, "image.png");

        // Create a simple text file that will be embedded as the initial OLE object.
        File.WriteAllText(textFilePath, "This is the original embedded text file.");

        // Create a minimal PNG image (1x1 pixel) without using System.Drawing.
        // The PNG data is a base64‑encoded transparent pixel.
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK2cAAAAASUVORK5CYII=");
        File.WriteAllBytes(imageFilePath, pngBytes);

        // -------------------------------------------------
        // Step 1: Create a document and embed the initial OLE object (the text file).
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph describing the original OLE object.
        builder.Writeln("Original OLE object (embedded text file):");

        // Embed the text file as an OLE object (not as an icon).
        using (FileStream txtStream = File.OpenRead(textFilePath))
        {
            // progId "Package" is a generic OLE container.
            builder.InsertOleObject(txtStream, "Package", false, null);
        }

        // Save the document that contains the original OLE object.
        doc.Save(originalDocPath);

        // -------------------------------------------------
        // Step 2: Load the document, remove the existing OLE object,
        // and replace it with a new image OLE object using InsertOleObject.
        // -------------------------------------------------
        Document loadedDoc = new Document(originalDocPath);
        DocumentBuilder replaceBuilder = new DocumentBuilder(loadedDoc);

        // Locate the first shape that is an OLE object.
        Shape oleShape = loadedDoc.GetChildNodes(NodeType.Shape, true)
                                  .OfType<Shape>()
                                  .FirstOrDefault(s => s.ShapeType == ShapeType.OleObject);

        if (oleShape != null)
        {
            // Remove the existing OLE shape.
            oleShape.Remove();
        }

        // Move the builder to the end of the document to insert the new OLE object.
        replaceBuilder.MoveToDocumentEnd();
        replaceBuilder.Writeln("\nReplaced OLE object (embedded image as icon):");

        // Insert the image as an OLE object, displayed as an icon.
        // progId "Package" treats the embedded file as a generic package.
        using (FileStream imageDataStream = File.OpenRead(imageFilePath))
        {
            // No separate presentation stream is required; passing null lets Aspose.Words use a default icon.
            replaceBuilder.InsertOleObject(imageDataStream, "Package", true, null);
        }

        // Save the modified document.
        loadedDoc.Save(modifiedDocPath);

        // Clean up temporary files (optional).
        // Comment out the following lines if you want to inspect the generated documents.
        //File.Delete(textFilePath);
        //File.Delete(imageFilePath);
        //File.Delete(originalDocPath);
        //File.Delete(modifiedDocPath);
        //Directory.Delete(tempFolder);
    }
}
