using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertOleObjectExample
{
    static void Main()
    {
        // Path to the folder that contains the source files.
        string dataDir = @"C:\Data\";

        // Path to the OLE source file (e.g., an Excel workbook) that will be embedded.
        string oleFilePath = Path.Combine(dataDir, "SampleSpreadsheet.xlsx");

        // Create a new DOCM document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the newly created document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a description before the OLE object.
        builder.Writeln("Embedded Excel spreadsheet:");

        // Insert the OLE object.
        // Parameters:
        //   fileName   – full path to the source file.
        //   isLinked   – false to embed the object (true would create a link).
        //   asIcon     – false to display the object content; true would show it as an icon.
        //   presentation – null to use the default presentation image.
        builder.InsertOleObject(oleFilePath, false, false, null);

        // Save the document as a DOCM file.
        string outputPath = Path.Combine(dataDir, "DocumentWithOleObject.docm");
        doc.Save(outputPath);
    }
}
