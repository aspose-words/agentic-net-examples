using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertOleObjectExample
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Path to the file that will be embedded as an OLE object.
        // In this example we embed a ZIP archive, but any file type can be used.
        string oleFilePath = @"C:\Data\SampleArchive.zip";

        // Read the file bytes into a memory stream.
        byte[] oleBytes = File.ReadAllBytes(oleFilePath);
        using (MemoryStream oleStream = new MemoryStream(oleBytes))
        {
            // Insert the OLE object as an icon.
            // Parameters:
            //   stream   – the data stream of the file to embed.
            //   progId   – "Package" indicates a generic OLE package.
            //   asIcon   – true to display the object as an icon.
            //   presentation – null to use the default icon.
            Shape oleShape = builder.InsertOleObject(oleStream, "Package", true, null);

            // Access the underlying OLE package to set display properties.
            OlePackage package = oleShape.OleFormat.OlePackage;
            package.FileName = Path.GetFileName(oleFilePath);      // File name shown when opened.
            package.DisplayName = "Embedded Sample Archive";      // Caption displayed under the icon.
        }

        // Save the document to a DOCX file.
        string outputPath = @"C:\Output\OleObjectDocument.docx";
        doc.Save(outputPath);
    }
}
