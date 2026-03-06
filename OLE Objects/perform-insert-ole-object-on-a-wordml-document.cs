using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertOleObjectExample
{
    static void Main()
    {
        // Path to the folder that contains the source file and where the result will be saved.
        string dataDir = @"C:\Data\";

        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Load the file that will be embedded as an OLE object.
        // In this example we embed a ZIP archive using the generic "Package" ProgID.
        byte[] zipBytes = File.ReadAllBytes(Path.Combine(dataDir, "cat001.zip"));
        using (MemoryStream zipStream = new MemoryStream(zipBytes))
        {
            // Insert the OLE object as an icon (asIcon = true) without a custom presentation image.
            // Parameters: stream, progId, asIcon, presentation
            Shape oleShape = builder.InsertOleObject(zipStream, "Package", true, null);

            // Set the file name and display name that will appear when the OLE object is opened.
            oleShape.OleFormat.OlePackage.FileName = "cat001.zip";
            oleShape.OleFormat.OlePackage.DisplayName = "Sample ZIP Archive";
        }

        // Save the document to a DOCX file.
        doc.Save(Path.Combine(dataDir, "InsertOleObject.docx"));
    }
}
