using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertOleObjectExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a heading for the OLE object.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Embedded OLE Object");

        // Return to normal paragraph style for the OLE insertion.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;

        // Path to the file that will be embedded as an OLE object.
        // Replace with an actual file path on your system.
        string oleFilePath = @"C:\Path\To\Your\File.xlsx";

        // Optional: path to an icon image to display for the OLE object.
        // If null, Aspose.Words will use a predefined icon.
        string iconPath = @"C:\Path\To\Your\Icon.ico";

        // Insert the OLE object as an icon (embedded, not linked).
        // Parameters: file name, isLinked = false, asIcon = true, presentation = icon stream.
        using (FileStream iconStream = File.OpenRead(iconPath))
        {
            builder.InsertOleObject(oleFilePath, false, true, iconStream);
        }

        // Save the document to a DOCX file.
        // Replace with your desired output path.
        string outputPath = @"C:\Path\To\Output\OleObjectDocument.docx";
        doc.Save(outputPath);
    }
}
