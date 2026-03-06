using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load the existing DOCX document
        Document doc = new Document("InputDocument.docx");

        // Find the first OLE object shape in the document
        Shape oleShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
        if (oleShape == null || oleShape.ShapeType != ShapeType.OleObject)
            throw new InvalidOperationException("No OLE object found in the document.");

        // Position the builder at the OLE shape so we can replace it
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveTo(oleShape);

        // Remove the old OLE shape
        oleShape.Remove();

        // Insert a new OLE object (for example, an Excel spreadsheet) in its place
        using (FileStream newOleStream = new FileStream("NewSpreadsheet.xlsx", FileMode.Open, FileAccess.Read))
        {
            // Insert the OLE object from a stream.
            // Parameters: stream, progId, asIcon (false = display content), presentation (null = default icon)
            builder.InsertOleObject(newOleStream, "Excel.Sheet", false, null);
        }

        // Save the updated document
        doc.Save("UpdatedDocument.docx");
    }
}
