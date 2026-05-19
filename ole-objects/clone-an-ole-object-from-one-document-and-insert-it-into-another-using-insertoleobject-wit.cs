using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare a simple file to embed as an OLE object.
        string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(dataDir);
        string embeddedFilePath = Path.Combine(dataDir, "Sample.txt");
        File.WriteAllText(embeddedFilePath, "This is sample text for OLE embedding.");

        // Create the source document and embed the OLE object.
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);
        using (FileStream embedStream = File.OpenRead(embeddedFilePath))
        {
            // Insert as an OLE Package (generic file).
            sourceBuilder.InsertOleObject(embedStream, "Package", false, null);
        }
        string sourcePath = Path.Combine(dataDir, "Source.docx");
        sourceDoc.Save(sourcePath);

        // Load the source document and extract the OLE object's data.
        Document loadedSource = new Document(sourcePath);
        Shape oleShape = loadedSource.GetChildNodes(NodeType.Shape, true)
            .OfType<Shape>()
            .FirstOrDefault(s => s.ShapeType == ShapeType.OleObject);
        if (oleShape == null)
            throw new InvalidOperationException("No OLE object found in the source document.");

        OleFormat oleFormat = oleShape.OleFormat;
        string progId = oleFormat.ProgId; // e.g., "Package"

        // Save the OLE data into a memory stream.
        using (MemoryStream oleDataStream = new MemoryStream())
        {
            oleFormat.Save(oleDataStream);
            oleDataStream.Position = 0; // Reset stream position for reading.

            // Create the destination document and insert the cloned OLE object.
            Document destDoc = new Document();
            DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
            destBuilder.Writeln("Cloned OLE object inserted below:");
            destBuilder.InsertOleObject(oleDataStream, progId, false, null);
            string destPath = Path.Combine(dataDir, "Destination.docx");
            destDoc.Save(destPath);
        }
    }
}
