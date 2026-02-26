using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new DOT (Word template) document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph to describe the OLE object.
        builder.Writeln("Embedded ZIP package OLE object:");

        // Load the file that will be embedded as an OLE package.
        // Replace the path with the actual location of your file.
        byte[] zipBytes = File.ReadAllBytes("cat001.zip");

        using (MemoryStream zipStream = new MemoryStream(zipBytes))
        {
            // Insert the OLE object from the stream.
            // Parameters: stream, progId ("Package" for generic OLE package), asIcon = true, presentation = null.
            Shape oleShape = builder.InsertOleObject(zipStream, "Package", true, null);

            // Set the file name and display name that Word will show for the OLE object.
            oleShape.OleFormat.OlePackage.FileName = "cat001.zip";
            oleShape.OleFormat.OlePackage.DisplayName = "cat001.zip";
        }

        // Save the document as a DOT template.
        doc.Save("TemplateWithOle.dot");
    }
}
