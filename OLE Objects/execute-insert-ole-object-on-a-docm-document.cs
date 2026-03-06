using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertOleObjectExample
{
    static void Main()
    {
        // Create a new empty DOCM document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Path to the file that will be embedded as an OLE object.
        // In this example we embed a ZIP archive, but any file type can be used.
        string oleFilePath = @"C:\Data\sample.zip";

        // Read the file into a memory stream.
        byte[] fileBytes = File.ReadAllBytes(oleFilePath);
        using (MemoryStream oleStream = new MemoryStream(fileBytes))
        {
            // Insert the OLE object from the stream.
            // Parameters:
            //   stream   – the data stream of the file.
            //   progId   – the ProgID that identifies the OLE type ("Package" for generic files).
            //   asIcon   – true to display the object as an icon, false to display its content.
            //   presentation – null to use the default icon.
            Shape oleShape = builder.InsertOleObject(oleStream, "Package", true, null);

            // Optionally customize the icon caption and the embedded package properties.
            oleShape.OleFormat.OlePackage.FileName = Path.GetFileName(oleFilePath);
            oleShape.OleFormat.OlePackage.DisplayName = "Embedded ZIP archive";
        }

        // Save the document as a DOCM file (macro-enabled Word document).
        string outputPath = @"C:\Output\DocumentWithOleObject.docm";
        doc.Save(outputPath);
    }
}
