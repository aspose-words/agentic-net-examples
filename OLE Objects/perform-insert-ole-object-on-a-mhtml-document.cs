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

        // Initialize a DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph describing the OLE object.
        builder.Writeln("Embedded OLE object:");

        // Path to the file that will be embedded as an OLE object.
        // In this example we use a generic package (e.g., a ZIP file) so the ProgId is "Package".
        string oleFilePath = "sample.zip";

        // Open the file as a stream and insert it as an OLE object.
        using (FileStream oleStream = new FileStream(oleFilePath, FileMode.Open, FileAccess.Read))
        {
            // InsertOleObject(stream, progId, asIcon, presentation)
            // progId: "Package" – generic OLE package.
            // asIcon: false – display the object content (if supported) rather than an icon.
            // presentation: null – let Aspose.Words choose a default presentation image.
            builder.InsertOleObject(oleStream, "Package", false, null);
        }

        // Save the document in MHTML format.
        doc.Save("Output.mhtml", SaveFormat.Mhtml);
    }
}
