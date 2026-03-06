using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertOleIntoTxt
{
    static void Main()
    {
        // Load an existing TXT file into a Word document.
        Document doc = new Document("Input.txt");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a description before the OLE object.
        builder.Writeln("Embedded Excel spreadsheet:");

        // Open the file that will be embedded as an OLE object.
        using (FileStream excelStream = new FileStream("Spreadsheet.xlsx", FileMode.Open, FileAccess.Read))
        {
            // Insert the OLE object.
            // Parameters: stream, progId, asIcon (false = show content), presentation (null = default icon).
            Shape oleShape = builder.InsertOleObject(excelStream, "Excel.Sheet", false, null);

            // Optional: set the size of the inserted shape.
            oleShape.Width = 400;
            oleShape.Height = 300;
        }

        // Save the result. OLE objects cannot be represented in plain TXT,
        // so we save the document in a format that supports them (DOCX).
        doc.Save("Result.docx");
    }
}
