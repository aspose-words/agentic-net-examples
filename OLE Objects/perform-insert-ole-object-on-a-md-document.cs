using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertOleObjectExample
{
    static void Main()
    {
        // Path to the source OLE file (e.g., an Excel workbook).
        string oleFilePath = @"C:\Data\Sample.xlsx";

        // Optional: path to an image that will be used as the OLE icon.
        // If you set asIcon to false, this parameter can be null.
        string iconPath = @"C:\Data\Icon.png";

        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a description before the OLE object.
        builder.Writeln("Below is an embedded Excel workbook as an OLE object:");

        // Open the optional icon image as a stream (null if not needed).
        Stream iconStream = null;
        if (File.Exists(iconPath))
        {
            iconStream = new FileStream(iconPath, FileMode.Open, FileAccess.Read);
        }

        // Insert the OLE object.
        // Parameters:
        //   fileName   – full path to the OLE file.
        //   isLinked   – false = embed the object, true = create a link.
        //   asIcon     – true = display as an icon, false = display the object's content.
        //   presentation – stream containing the icon image (null to use default icon).
        builder.InsertOleObject(oleFilePath, false, true, iconStream);

        // Clean up the icon stream if it was opened.
        iconStream?.Dispose();

        // Save the document to disk.
        string outputPath = @"C:\Output\OleObjectDocument.docx";
        doc.Save(outputPath);
    }
}
