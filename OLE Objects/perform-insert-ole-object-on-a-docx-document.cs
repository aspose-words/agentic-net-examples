using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertOleObjectExample
{
    static void Main()
    {
        // Paths to input and output folders – adjust as needed.
        string myDir = @"C:\Data\";
        string artifactsDir = @"C:\Output\";

        // Load the binary content of a ZIP file that will be embedded as an OLE package.
        byte[] zipBytes = File.ReadAllBytes(Path.Combine(myDir, "cat001.zip"));
        using (MemoryStream zipStream = new MemoryStream(zipBytes))
        {
            // Create a new blank document and a DocumentBuilder to edit it.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a descriptive paragraph before the OLE object.
            builder.Writeln("Embedded ZIP package:");

            // Insert the OLE object from the stream.
            // Parameters: stream, progId ("Package" for generic OLE package), asIcon = true, presentation = null.
            Shape oleShape = builder.InsertOleObject(zipStream, "Package", true, null);

            // Set the file name and display name that will appear when the OLE object is opened.
            oleShape.OleFormat.OlePackage.FileName = "cat001.zip";
            oleShape.OleFormat.OlePackage.DisplayName = "cat001.zip";

            // Save the resulting document.
            doc.Save(Path.Combine(artifactsDir, "InsertOleObject.docx"));
        }
    }
}
