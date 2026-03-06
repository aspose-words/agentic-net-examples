using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertOleObjectExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Path to the file that will be embedded as an OLE object.
        string oleFilePath = @"C:\Temp\Spreadsheet.xlsx";

        // Open the file as a stream.
        using (Stream oleStream = File.Open(oleFilePath, FileMode.Open, FileAccess.Read))
        {
            // Insert the OLE object.
            // progId "Package" is a generic identifier for unknown file types.
            // asIcon = false embeds the object content directly (no icon).
            // presentation = null lets Aspose.Words choose a default presentation.
            Shape oleShape = builder.InsertOleObject(oleStream, "Package", false, null);
        }

        // Save the document to a DOCX file.
        doc.Save(@"C:\Temp\OleObjectDocument.docx");
    }
}
