using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertOleObjectExample
{
    static void Main()
    {
        // Path to the file that will be embedded as an OLE object.
        string oleSourcePath = @"C:\Data\Sample.xlsx";

        // Path where the resulting RTF document will be saved.
        string outputPath = @"C:\Output\DocumentWithOle.rtf";

        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a description before the OLE object.
        builder.Writeln("Embedded Excel spreadsheet:");

        // Open the source file as a stream.
        using (Stream oleStream = File.OpenRead(oleSourcePath))
        {
            // Insert the OLE object from the stream.
            // Parameters:
            //   oleStream   – stream containing the file data.
            //   "Package"   – ProgID for generic OLE package.
            //   true        – display as an icon.
            //   null        – use default icon image.
            Shape oleShape = builder.InsertOleObject(oleStream, "Package", true, null);

            // Optionally set a custom file name and display name for the OLE package.
            oleShape.OleFormat.OlePackage.FileName = Path.GetFileName(oleSourcePath);
            oleShape.OleFormat.OlePackage.DisplayName = Path.GetFileNameWithoutExtension(oleSourcePath);
        }

        // Save the document in RTF format.
        doc.Save(outputPath, SaveFormat.Rtf);
    }
}
