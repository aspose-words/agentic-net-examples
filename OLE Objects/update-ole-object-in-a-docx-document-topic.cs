using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load an existing DOCX file that contains an OLE object.
        string inputPath = @"C:\Docs\InputDocument.docx";
        var doc = new Document(inputPath);

        // Use a DocumentBuilder to navigate and edit the document.
        var builder = new DocumentBuilder(doc);

        // Locate the first shape that holds an OLE object.
        var oleShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
        if (oleShape == null || oleShape.OleFormat == null)
            throw new InvalidOperationException("No OLE object found in the document.");

        // ---------------------------------------------------------------------
        // Example 1: Update OLE object properties (e.g., enable auto‑update).
        // ---------------------------------------------------------------------
        oleShape.OleFormat.AutoUpdate = true; // Word will refresh the linked data automatically.
        // Note: IconCaption is read‑only in Aspose.Words and cannot be set directly.
        // If you need a custom caption, insert a new OLE object with the desired caption.

        // ---------------------------------------------------------------------
        // Example 2: Replace the embedded OLE data with a new file.
        // ---------------------------------------------------------------------
        // Move the builder cursor to the existing OLE shape.
        builder.MoveTo(oleShape);

        // Insert a new OLE object (e.g., a new Excel workbook) in place of the old one.
        // The InsertOleObject overload allows specifying an icon caption.
        using (FileStream newFileStream = File.OpenRead(@"C:\Data\NewSpreadsheet.xlsx"))
        {
            // "Excel.Sheet.12" is the ProgID for an Excel workbook.
            // asIcon = false displays the content directly; set to true to show an icon.
            // The last parameter is the icon caption (used only when asIcon is true).
            var newOleShape = builder.InsertOleObject(newFileStream, "Excel.Sheet.12", false, null);

            // Optionally copy formatting from the old shape to the new one.
            newOleShape.Width = oleShape.Width;
            newOleShape.Height = oleShape.Height;
        }

        // Remove the original OLE shape now that it has been replaced.
        oleShape.Remove();

        // ---------------------------------------------------------------------
        // Save the modified document.
        // ---------------------------------------------------------------------
        string outputPath = @"C:\Docs\UpdatedDocument.docx";
        doc.Save(outputPath);
    }
}
