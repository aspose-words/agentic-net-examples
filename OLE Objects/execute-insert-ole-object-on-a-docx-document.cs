using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertOleObjectExample
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Path to the file that will be embedded as an OLE object (e.g., an Excel workbook).
        string oleFilePath = @"C:\Data\Sample.xlsx";

        // Write a description before the OLE object.
        builder.Writeln("Embedded Excel workbook:");

        // Open the source file as a stream and insert it as an OLE object.
        using (FileStream oleStream = File.Open(oleFilePath, FileMode.Open, FileAccess.Read))
        {
            // progId identifies the OLE type. "Excel.Sheet" works for .xlsx files.
            // asIcon = false inserts the full content (not just an icon).
            // presentation = null lets Aspose use the default preview image.
            Shape oleShape = builder.InsertOleObject(oleStream, "Excel.Sheet", false, null);

            // Optionally resize the inserted OLE object.
            oleShape.Width = 400;   // points
            oleShape.Height = 300;  // points
        }

        // Save the resulting document as DOCX.
        doc.Save(@"C:\Output\OleObjectDocument.docx");
    }
}
