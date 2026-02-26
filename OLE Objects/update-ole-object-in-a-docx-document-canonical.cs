using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class UpdateOleObjectExample
{
    static void Main()
    {
        // Load the existing DOCX document.
        Document doc = new Document("InputDocument.docx");

        // Find the first shape that contains an OLE object.
        Shape oldOleShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
        if (oldOleShape == null || oldOleShape.OleFormat == null)
        {
            Console.WriteLine("No OLE object found in the document.");
            return;
        }

        // Create a DocumentBuilder positioned at the OLE shape.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveTo(oldOleShape);

        // Insert a new OLE object (replace the old one) from a file stream.
        // Adjust the file path, ProgID and other parameters as needed.
        using (FileStream newOleStream = File.Open("NewEmbeddedFile.xlsx", FileMode.Open, FileAccess.Read))
        {
            // Insert the new OLE object at the current builder position.
            // Parameters: stream, progId, asIcon (false = display content), presentation (null = default icon).
            builder.InsertOleObject(newOleStream, "Excel.Sheet", false, null);
        }

        // Remove the original OLE shape from the document.
        oldOleShape.Remove();

        // Save the updated document.
        doc.Save("UpdatedDocument.docx");
    }
}
