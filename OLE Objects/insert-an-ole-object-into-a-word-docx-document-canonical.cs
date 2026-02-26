using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertOleObjectExample
{
    static void Main()
    {
        // Path to the file that will be embedded as an OLE package (e.g., a ZIP archive)
        string zipFilePath = @"C:\Data\sample.zip";

        // Path where the resulting DOCX will be saved
        string outputPath = @"C:\Output\OleObject.docx";

        // Create a new blank Word document
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a caption before the OLE object
        builder.Writeln("Embedded ZIP archive as OLE object:");

        // Load the ZIP file into a memory stream
        using (MemoryStream zipStream = new MemoryStream(File.ReadAllBytes(zipFilePath)))
        {
            // Insert the OLE object from the stream.
            // Parameters:
            //   stream   – the data stream containing the file bytes
            //   progId   – "Package" indicates a generic OLE package
            //   asIcon   – true to display the object as an icon
            //   presentation – null to use the default icon
            Shape oleShape = builder.InsertOleObject(zipStream, "Package", true, null);

            // Set the file name and display name that Word will show when the object is opened
            oleShape.OleFormat.OlePackage.FileName = Path.GetFileName(zipFilePath);
            oleShape.OleFormat.OlePackage.DisplayName = "Sample Archive.zip";
        }

        // Save the document to the specified location
        doc.Save(outputPath);
    }
}
