using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class OleCloneExample
{
    static void Main()
    {
        // ---------- Create a source document with an OLE object ----------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Writeln("Source document with OLE object:");

        // Create a simple text file in memory to embed as OLE.
        byte[] oleFileBytes = System.Text.Encoding.UTF8.GetBytes("This is the embedded OLE content.");
        using (MemoryStream oleFileStream = new MemoryStream(oleFileBytes))
        {
            // Insert the OLE object. Use "Package" as a generic ProgID.
            srcBuilder.InsertOleObject(oleFileStream, "Package", false, null);
        }

        // Save the source document (optional, for inspection).
        srcDoc.Save("SourceWithOle.docx");

        // ---------- Locate the OLE shape in the source document ----------
        Shape srcOleShape = (Shape)srcDoc.GetChild(NodeType.Shape, 0, true);
        if (srcOleShape == null || srcOleShape.OleFormat == null)
            throw new InvalidOperationException("No OLE object found in the source document.");

        OleFormat srcOleFormat = srcOleShape.OleFormat;
        string progId = srcOleFormat.ProgId;

        // Save the embedded OLE data into a memory stream.
        using (MemoryStream oleDataStream = new MemoryStream())
        {
            srcOleFormat.Save(oleDataStream);
            oleDataStream.Position = 0; // Reset for reading.

            // ---------- Create a destination document ----------
            Document dstDoc = new Document();
            DocumentBuilder dstBuilder = new DocumentBuilder(dstDoc);
            dstBuilder.Writeln("Destination document before cloning OLE object:");
            dstBuilder.MoveToDocumentEnd();

            // Insert the cloned OLE object.
            dstBuilder.InsertOleObject(oleDataStream, progId, false, null);
            dstBuilder.Writeln("\nDestination document after cloning OLE object.");

            // Save the destination document.
            dstDoc.Save("DestinationWithClonedOle.docx");
        }

        Console.WriteLine("OLE object cloned successfully.");
    }
}
