using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertOleObjectExample
{
    static void Main()
    {
        // Define directories for input and output files.
        string dataDir = @"C:\Data\";
        string outputDir = @"C:\Output\";

        // Path to the file that will be embedded as an OLE object.
        // In this example we embed an Excel spreadsheet.
        string oleFilePath = Path.Combine(dataDir, "SampleSpreadsheet.xlsx");

        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a description before the OLE object.
        builder.Writeln("Embedded Excel spreadsheet:");

        // Insert the OLE object.
        // Parameters:
        //   fileName   – full path to the source file.
        //   isLinked   – false to embed the file (true would create a link).
        //   asIcon     – false to display the content; true would display an icon.
        //   presentation – null to use the default icon or preview image.
        builder.InsertOleObject(oleFilePath, false, false, null);

        // Save the resulting document.
        string outputPath = Path.Combine(outputDir, "DocumentWithOleObject.docx");
        doc.Save(outputPath);
    }
}
