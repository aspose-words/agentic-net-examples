using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class UpdateOleObjectExample
{
    static void Main()
    {
        // Load the existing DOCX document that contains an OLE object.
        Document doc = new Document("InputDocument.docx");

        // Find the first shape that holds an OLE object.
        Shape oleShape = doc.GetChild(NodeType.Shape, 0, true) as Shape;
        if (oleShape == null || oleShape.ShapeType != ShapeType.OleObject)
        {
            Console.WriteLine("No OLE object found in the document.");
            return;
        }

        // Access the OLE format of the shape.
        OleFormat oleFormat = oleShape.OleFormat;

        // 1. Enable automatic update for linked OLE objects.
        oleFormat.AutoUpdate = true;

        // NOTE: IconCaption is a read‑only property in Aspose.Words and cannot be set directly.
        // If you need a different caption you must replace the OLE object with a new one.

        // 2. If the OLE object is a package, you can modify its displayed file name.
        if (oleFormat.OlePackage != null)
        {
            oleFormat.OlePackage.FileName = "UpdatedFile.zip";
            oleFormat.OlePackage.DisplayName = "UpdatedFile.zip";
        }

        // 3. For linked objects, you can change the source file path.
        // oleFormat.SourceFullName = @"C:\NewPath\NewFile.xlsx";

        // Save the modified document.
        doc.Save("UpdatedDocument.docx");
    }
}
