using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Path to the file that will be embedded as an OLE object.
        // In this example we embed a ZIP archive, but any file can be used.
        string sourceFilePath = @"C:\Data\sample.zip";

        // Path where the resulting DOCX document will be saved.
        string outputFilePath = @"C:\Output\OleObject.docx";

        // Create a new blank Word document.
        Document doc = new Document();

        // Initialize a DocumentBuilder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a description before the OLE object.
        builder.Writeln("Embedded ZIP archive as OLE object:");

        // Load the source file into a memory stream.
        using (MemoryStream fileStream = new MemoryStream(File.ReadAllBytes(sourceFilePath)))
        {
            // Insert the OLE object from the stream.
            // Parameters:
            //   fileStream   – stream containing the file data.
            //   "Package"    – ProgID for generic OLE packages.
            //   true         – display the object as an icon.
            //   null         – use the default icon image.
            Shape oleShape = builder.InsertOleObject(fileStream, "Package", true, null);

            // Access the underlying OLE package to set a friendly file name and display name.
            oleShape.OleFormat.OlePackage.FileName = Path.GetFileName(sourceFilePath);
            oleShape.OleFormat.OlePackage.DisplayName = Path.GetFileName(sourceFilePath);
        }

        // Save the document in DOCX format.
        doc.Save(outputFilePath);
    }
}
