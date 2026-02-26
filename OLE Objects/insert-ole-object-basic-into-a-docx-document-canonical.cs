using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertOleObjectExample
{
    static void Main()
    {
        // Path to the file that will be embedded as an OLE object.
        // In this example we embed a ZIP archive, but any file can be used.
        string sourceFilePath = @"C:\Temp\myfile.zip";

        // Path where the resulting DOCX document will be saved.
        string outputDocPath = @"C:\Temp\output.docx";

        // Create a new blank Word document.
        Document doc = new Document();

        // Initialize a DocumentBuilder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Load the source file into a memory stream.
        // The stream will be passed to InsertOleObject.
        using (MemoryStream oleStream = new MemoryStream(File.ReadAllBytes(sourceFilePath)))
        {
            // Insert the OLE object.
            // Parameters:
            //   oleStream   – stream containing the file data.
            //   "Package"   – ProgID for a generic OLE package.
            //   true        – display the object as an icon.
            //   null        – no custom presentation image (default icon will be used).
            Shape oleShape = builder.InsertOleObject(oleStream, "Package", true, null);

            // Optionally set the file name and display name that appear when the OLE object is opened.
            oleShape.OleFormat.OlePackage.FileName = Path.GetFileName(sourceFilePath);
            oleShape.OleFormat.OlePackage.DisplayName = "Embedded ZIP Archive";
        }

        // Save the document to the specified path.
        doc.Save(outputDocPath);
    }
}
