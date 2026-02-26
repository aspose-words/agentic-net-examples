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
        builder.Writeln("Embedded OLE object (sample.txt):");

        // Path to the file that will be embedded as an OLE object.
        string filePath = "sample.txt";

        // Ensure the sample file exists; create it if necessary.
        if (!File.Exists(filePath))
        {
            File.WriteAllText(filePath, "This is sample text inside the OLE package.");
        }

        // Insert the OLE object as an icon.
        // Parameters:
        //   filePath   – full path to the source file.
        //   false      – embed the object (not a link).
        //   true       – display as an icon.
        //   null       – use the default icon provided by Aspose.Words.
        Shape oleShape = builder.InsertOleObject(filePath, false, true, null);

        // NOTE: In recent versions of Aspose.Words the IconCaption property is read‑only.
        // If you need a custom caption, you must create a separate caption paragraph
        // or use a custom icon image. The line below has been removed to fix the build error.
        // oleShape.OleFormat.IconCaption = "Sample Text File";

        // Save the document in RTF format.
        doc.Save("OleObjectDocument.rtf");
    }
}
