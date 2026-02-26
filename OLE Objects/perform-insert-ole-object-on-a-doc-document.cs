using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Ole;

class Program
{
    static void Main()
    {
        // Path to the file that will be embedded as an OLE object.
        // In this example we embed a ZIP archive, but any file can be used.
        string zipFilePath = @"C:\Data\sample.zip";

        // Path where the resulting document will be saved.
        string outputDocPath = @"C:\Output\OleObject.docx";

        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a description before the OLE object.
        builder.Writeln("Embedded ZIP archive as OLE object:");

        // Read the file into a memory stream.
        using (FileStream fileStream = new FileStream(zipFilePath, FileMode.Open, FileAccess.Read))
        {
            // Insert the OLE object from the stream.
            // Parameters:
            //   stream   – the data stream of the file.
            //   progId   – "Package" indicates a generic OLE package.
            //   asIcon   – true to display the object as an icon.
            //   presentation – null to use the default icon.
            Shape oleShape = builder.InsertOleObject(fileStream, "Package", true, null);

            // Access the OLE package to set a friendly file name and display name.
            OlePackage olePackage = oleShape.OleFormat.OlePackage;
            olePackage.FileName = Path.GetFileName(zipFilePath);      // e.g., "sample.zip"
            olePackage.DisplayName = "Sample ZIP Archive";
        }

        // Save the document to the specified location.
        doc.Save(outputDocPath);
    }
}
