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

        // Add a description before the OLE object.
        builder.Writeln("Embedded PDF file as an OLE object:");

        // Path to the file that will be embedded.
        string filePath = @"C:\Data\sample.pdf";

        // Open the file as a stream and insert it as an OLE object.
        using (Stream fileStream = File.Open(filePath, FileMode.Open, FileAccess.Read))
        {
            // Insert the OLE object using the generic "Package" progId.
            // Set asIcon to true to display the object as an icon.
            // Pass null for the presentation stream to use the default icon.
            builder.InsertOleObject(fileStream, "Package", true, null);
        }

        // Save the resulting document.
        doc.Save(@"C:\Output\OleObjectDocument.docx");
    }
}
