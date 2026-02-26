using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Path to the file that will be embedded as an OLE object.
        string oleFilePath = @"C:\Data\SamplePackage.zip";

        // Open the file as a stream.
        using (FileStream fileStream = new FileStream(oleFilePath, FileMode.Open, FileAccess.Read))
        {
            // Insert the OLE object. 
            // progId "Package" indicates a generic OLE package.
            // asIcon = true displays the object as an icon.
            // presentation = null lets Aspose.Words choose a default icon.
            Shape oleShape = builder.InsertOleObject(fileStream, "Package", true, null);

            // Set the file name and display name that appear when the OLE object is opened.
            oleShape.OleFormat.OlePackage.FileName = Path.GetFileName(oleFilePath);
            oleShape.OleFormat.OlePackage.DisplayName = Path.GetFileNameWithoutExtension(oleFilePath);
        }

        // Save the document to a DOCX file.
        doc.Save(@"C:\Output\OleObjectDocument.docx");
    }
}
