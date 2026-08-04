using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a source document and insert an OLE package object.
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);

        // Sample data to embed in the OLE package.
        byte[] sampleData = Encoding.UTF8.GetBytes("Hello from OLE package");
        using (MemoryStream dataStream = new MemoryStream(sampleData))
        {
            // Insert the OLE object as a package.
            Shape oleShape = srcBuilder.InsertOleObject(dataStream, "Package", false, null);
            // Set package metadata.
            oleShape.OleFormat.OlePackage.FileName = "sample.txt";
            oleShape.OleFormat.OlePackage.DisplayName = "Sample Text";
        }

        // Clone the source document (deep copy).
        Document clonedDoc = srcDoc.Clone();

        // Retrieve the OLE shape from the cloned document.
        Shape clonedShape = (Shape)clonedDoc.GetChild(NodeType.Shape, 0, true);
        OleFormat clonedOle = clonedShape.OleFormat;

        // Extract the OLE data into a memory stream.
        MemoryStream extractedStream = new MemoryStream();
        clonedOle.Save(extractedStream);
        extractedStream.Position = 0; // Reset stream position for reading.

        // Create a target document where the cloned OLE object will be inserted.
        Document targetDoc = new Document();
        DocumentBuilder targetBuilder = new DocumentBuilder(targetDoc);
        targetBuilder.Writeln("Cloned OLE object inserted below:");

        // Insert the extracted OLE data into the target document.
        targetBuilder.InsertOleObject(extractedStream, clonedOle.ProgId, false, null);

        // Save the resulting document.
        targetDoc.Save("ClonedOleObject.docx");
    }
}
