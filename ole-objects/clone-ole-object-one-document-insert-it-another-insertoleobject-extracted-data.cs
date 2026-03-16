using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class OleCloneExample
{
    static void Main()
    {
        // Load the source document that contains the OLE object.
        Document srcDoc = new Document("SourceWithOle.docx");

        // Locate the first shape that holds an OLE object.
        Shape srcOleShape = (Shape)srcDoc.GetChild(NodeType.Shape, 0, true);
        if (srcOleShape == null || srcOleShape.OleFormat == null)
            throw new InvalidOperationException("No OLE object found in the source document.");

        // Access the OLE format of the shape.
        OleFormat srcOleFormat = srcOleShape.OleFormat;

        // Retrieve the ProgID of the OLE object (e.g., "Excel.Sheet.12").
        string progId = srcOleFormat.ProgId;

        // Save the embedded OLE data into a memory stream.
        using (MemoryStream oleDataStream = new MemoryStream())
        {
            srcOleFormat.Save(oleDataStream);
            oleDataStream.Position = 0; // Reset stream position for reading.

            // Load the destination document where the OLE object will be cloned.
            Document dstDoc = new Document("Destination.docx");
            DocumentBuilder builder = new DocumentBuilder(dstDoc);

            // Move the cursor to the desired insertion point (here, the end of the document).
            builder.MoveToDocumentEnd();

            // Insert the OLE object using the extracted data stream.
            // asIcon = false – insert the object in its normal (content) view.
            // presentation = null – let Aspose.Words choose a default icon if needed.
            builder.InsertOleObject(oleDataStream, progId, false, null);

            // Save the modified destination document.
            dstDoc.Save("DestinationWithClonedOle.docx");
        }
    }
}
