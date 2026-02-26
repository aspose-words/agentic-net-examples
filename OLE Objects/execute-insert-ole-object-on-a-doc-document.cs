using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Path to the file that will be embedded as an OLE object.
        // Example: an Excel workbook.
        string dataDir = @"C:\Data\";
        string oleFilePath = Path.Combine(dataDir, "Sample.xlsx");

        // Write a caption before the OLE object.
        builder.Writeln("Embedded Excel OLE object:");

        // Open the source file as a stream and insert it as an OLE object.
        // ProgId "Excel.Sheet" tells Word the type of the embedded object.
        // The third argument (true) displays the object as an icon.
        // The fourth argument (null) lets Aspose.Words choose a default icon image.
        using (FileStream fileStream = new FileStream(oleFilePath, FileMode.Open, FileAccess.Read))
        {
            builder.InsertOleObject(fileStream, "Excel.Sheet", true, null);
        }

        // Save the document in the legacy DOC format.
        string outPath = Path.Combine(dataDir, "OleObjectDocument.doc");
        doc.Save(outPath);
    }
}
