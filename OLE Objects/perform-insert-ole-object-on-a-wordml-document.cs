using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertOleObjectExample
{
    static void Main()
    {
        // Load an existing WORDML document.
        // Replace "InputWordML.xml" with the path to your WORDML file.
        Document doc = new Document("InputWordML.xml");

        // Create a DocumentBuilder to modify the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a description before the OLE object.
        builder.Writeln("Embedded OLE object (ZIP package):");

        // Path to the file that will be embedded as an OLE object.
        // In this example we embed a ZIP archive, but any file type can be used.
        string oleFilePath = "cat001.zip";

        // Insert the OLE object.
        // Parameters:
        //   fileName   – full path to the file to embed.
        //   isLinked   – false to embed the file (true would create a link).
        //   asIcon     – false to display the object content; true to display as an icon.
        //   presentation – null to use the default icon or preview image.
        builder.InsertOleObject(oleFilePath, false, false, null);

        // Optionally, add a line break after the inserted object.
        builder.Writeln();

        // Save the modified document back to WORDML format.
        // Replace "OutputWordML.xml" with the desired output path.
        doc.Save("OutputWordML.xml");
    }
}
