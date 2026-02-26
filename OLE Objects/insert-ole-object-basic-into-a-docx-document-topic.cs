using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Initialize a DocumentBuilder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Path to the file that will be embedded as an OLE object.
        // The file can be any type that has an associated OLE handler (e.g., Excel, PowerPoint, etc.).
        string oleFilePath = "Spreadsheet.xlsx";

        // Insert the OLE object:
        //   - isLinked = false  -> embed the file data into the document.
        //   - asIcon   = false  -> display the actual content, not an icon.
        //   - presentation = null -> let Aspose.Words choose the default presentation.
        builder.InsertOleObject(oleFilePath, false, false, null);

        // Save the resulting document to a DOCX file.
        doc.Save("InsertOleObject.docx");
    }
}
