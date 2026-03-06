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

        // Initialize a DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Path to the file that will be embedded as an OLE object.
        // Replace with the actual path to your file (e.g., a ZIP archive, Excel file, etc.).
        string filePath = @"C:\Data\sample.zip";

        // Read the file into a byte array.
        byte[] fileBytes = File.ReadAllBytes(filePath);

        // Insert the OLE object from a memory stream.
        // Parameters:
        //   stream      – the data stream of the file.
        //   progId      – "Package" indicates a generic OLE package.
        //   asIcon      – true to display the object as an icon.
        //   presentation– null to use the default icon provided by Aspose.Words.
        using (MemoryStream stream = new MemoryStream(fileBytes))
        {
            Shape oleShape = builder.InsertOleObject(stream, "Package", true, null);

            // Optionally set a custom file name and display name for the embedded package.
            oleShape.OleFormat.OlePackage.FileName = "sample.zip";
            oleShape.OleFormat.OlePackage.DisplayName = "Sample ZIP Archive";
        }

        // Save the document to a DOCX file.
        // Replace with your desired output path.
        doc.Save(@"C:\Output\OleObjectExample.docx");
    }
}
